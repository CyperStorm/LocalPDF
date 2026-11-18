"""LocalPDF desktop interface and conversion engine."""
import importlib.util
import os
from pathlib import Path
import queue
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
PDF = ('*.pdf',)
IMAGES = ('*.jpg', '*.jpeg', '*.png', '*.webp', '*.bmp', '*.tif', '*.tiff')
TOOLS = {
    'Images to PDF': (IMAGES, '.pdf', True, '', 'Combine images into one PDF in the order shown. Multi-frame TIFF files include every frame.'),
    'PDF to JPG / PNG': (PDF, '.zip', False, 'Image format and resolution', 'Export one image per page. Images are delivered in a ZIP file.'),
    'Word to PDF': (('*.docx',), '.pdf', False, '', 'Uses desktop Microsoft Word on this PC.'),
    'Excel to PDF': (('*.xlsx',), '.pdf', False, '', 'Uses desktop Microsoft Excel and the workbook’s print settings.'),
    'PowerPoint to PDF': (('*.pptx',), '.pdf', False, '', 'Uses desktop Microsoft PowerPoint on this PC.'),
    'PDF to Word': (PDF, '.docx', False, '', 'Extracts selectable text into an editable document. Page breaks are retained. Requires extras.'),
    'PDF to Excel': (PDF, '.xlsx', False, '', 'Extracts detected tables to individual sheets. Works best with selectable text and visible table borders. Requires extras.'),
    'PDF to PowerPoint': (PDF, '.pptx', False, '', 'Places each PDF page on a slide as an image. Requires extras.'),
    'PDF to Text': (PDF, '.txt', False, '', 'Extracts selectable text into a UTF-8 file. Requires extras.'),
    'Merge PDFs': (PDF, '.pdf', True, '', 'Combine PDFs in the order shown. Requires extras.'),
    'Split PDF': (PDF, '.zip', False, '', 'Save each PDF page separately in a ZIP file. Requires extras.'),
    'Extract / reorder pages': (PDF, '.pdf', False, 'Page order', 'Enter page numbers and ranges such as 3, 1-2, 5. Requires extras.'),
    'Remove pages': (PDF, '.pdf', False, 'Pages to remove', 'Enter page numbers and ranges such as 2, 4-6. Requires extras.'),
    'Rotate PDF': (PDF, '.pdf', False, 'Clockwise rotation', 'Rotate every page. Requires extras.'),
    'Compress PDF': (PDF, '.pdf', False, '', 'Lossless optimization. Size reduction depends on the original PDF. Requires extras.'),
    'Protect PDF': (PDF, '.pdf', False, 'New password', 'Encrypt the PDF with AES-256. Requires extras.'),
    'Unlock PDF': (PDF, '.pdf', False, 'Current password', 'Remove encryption using the owner password. Requires extras.'),
}
OFFICE = {'Word to PDF': 'word', 'Excel to PDF': 'excel', 'PowerPoint to PDF': 'powerpoint'}
CHOICES = {
    'PDF to JPG / PNG': ('JPG · 150 DPI', 'JPG · 300 DPI', 'PNG · 150 DPI', 'PNG · 300 DPI'),
    'Rotate PDF': ('90', '180', '270'),
}
FIXED = {'Images to PDF', 'PDF to JPG / PNG', 'PDF to PowerPoint', *OFFICE}
POWERSHELL = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'


def run_script(name, *args):
    result = subprocess.run(
        [str(POWERSHELL), '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
         '-File', str(ROOT / name), *map(str, args)],
        capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if result.returncode:
        detail = result.stderr.strip() or result.stdout.strip()
        raise ValueError(f'{name} could not complete the conversion.\n\n{detail[:1200]}')


def check_office_file(path):
    # Block external content relationships and workbook data connections before Office opens the file.
    with zipfile.ZipFile(path) as package:
        for entry in package.namelist():
            if entry.startswith(('xl/externalLinks/', 'xl/queryTables/')) or entry == 'xl/connections.xml':
                raise ValueError('This workbook contains external data connections. Remove them before converting.')
            if entry.endswith('.rels'):
                for link in ET.fromstring(package.read(entry)):
                    if link.get('TargetMode', '').lower() == 'external' and not link.get('Type', '').endswith('/hyperlink'):
                        raise ValueError('This document references external content. Remove the links before converting.')


def images_to_pdf(paths, target):
    from PIL import Image, ImageOps, ImageSequence
    pages = []
    try:
        for path in paths:
            with Image.open(path) as image:
                for frame in ImageSequence.Iterator(image):
                    rgba = ImageOps.exif_transpose(frame).convert('RGBA')
                    page = Image.new('RGB', rgba.size, 'white')
                    page.paste(rgba, mask=rgba.getchannel('A'))
                    rgba.close()
                    pages.append(page)
        pages[0].save(target, 'PDF', save_all=True, append_images=pages[1:], resolution=150)
    finally:
        for page in pages:
            page.close()


def render_pdf(source, work, dpi):
    # Use a folder next to the staged output to keep page files local.
    pages = work / 'pages'
    pages.mkdir()
    run_script('Render.ps1', '-Source', source, '-Destination', pages, '-Dpi', dpi)
    result = sorted(pages.glob('page-*.png'))
    if not result:
        raise ValueError('The PDF contains no renderable pages.')
    return result


def pdf_to_images(source, target, option):
    from PIL import Image
    dpi = 300 if '300' in option else 150
    jpeg = option.startswith('JPG')
    pages = render_pdf(source, target.parent, dpi)
    with zipfile.ZipFile(target, 'w') as archive:
        for page in pages:
            size = tuple(map(int, page.with_suffix('.png.size').read_text().split()))
            output = page.with_name(page.stem + ('-export.jpg' if jpeg else '-export.png'))
            with Image.open(page) as image:
                # Windows may scale the render by the current display's scaling factor.
                rendered = image.resize(size, Image.Resampling.LANCZOS) if image.size != size else image
                if jpeg:
                    rendered.convert('RGB').save(output, quality=90, dpi=(dpi, dpi))
                else:
                    rendered.save(output, dpi=(dpi, dpi))
            archive.write(output, page.stem + output.suffix)


def pdf_to_slides(source, target):
    from PIL import Image
    from pptx import Presentation
    presentation = Presentation()
    for page in render_pdf(source, target.parent, 150):
        with Image.open(page) as image:
            width, height = image.size
        scale = min(presentation.slide_width / width, presentation.slide_height / height)
        w, h = round(width * scale), round(height * scale)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        slide.shapes.add_picture(str(page), (presentation.slide_width - w) // 2,
                                 (presentation.slide_height - h) // 2, w, h)
    presentation.save(str(target))


def page_order(value, count):
    if not value.strip():
        raise ValueError('Enter at least one page number.')
    pages = []
    for token in value.split(','):
        match = re.fullmatch(r'\s*([0-9]+)(?:\s*-\s*([0-9]+))?\s*', token)
        if not match:
            raise ValueError('Use comma-separated page numbers and ranges, such as 1, 3-5.')
        start = int(match[1])
        end = int(match[2] or match[1])
        if not (1 <= start <= count and 1 <= end <= count):
            raise ValueError(f'Page numbers must be between 1 and {count}.')
        step = 1 if start <= end else -1
        pages.extend(range(start - 1, end - 1 + step, step))
    return pages


def save_pdf(pdf, target, **options):
    pdf.save(str(target), garbage=4, deflate=True, **options)


def export_text(pdf, target, as_word):
    pages = [page.get_text(sort=True) for page in pdf]
    if not any(page.strip() for page in pages):
        raise ValueError('No selectable text was found. Scanned pages require OCR.')
    if not as_word:
        target.write_text('\n\f\n'.join(pages), encoding='utf-8')
        return
    from docx import Document
    document = Document()
    for index, text in enumerate(pages):
        if index:
            document.add_page_break()
        for line in text.splitlines():
            document.add_paragraph(re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', line))
    document.save(str(target))


def export_tables(pdf, target):
    from openpyxl import Workbook
    workbook = Workbook()
    workbook.remove(workbook.active)
    for number, page in enumerate(pdf, 1):
        for index, table in enumerate(page.find_tables().tables, 1):
            sheet = workbook.create_sheet(f'Page {number} Table {index}')
            for row_number, row in enumerate(table.extract(), 1):
                for column, value in enumerate(row, 1):
                    if value is not None:
                        cell = sheet.cell(row_number, column, re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', value))
                        cell.data_type = 's'
    if not workbook.sheetnames:
        raise ValueError('No tables were detected. Try a PDF with selectable text and visible table borders.')
    workbook.save(target)


def edit_pdf(tool, sources, target, option):
    import pymupdf as fitz
    if tool == 'Merge PDFs':
        with fitz.open() as merged:
            for path in sources:
                with fitz.open(str(path)) as pdf:
                    if pdf.needs_pass or pdf.metadata.get('encryption'):
                        raise ValueError(f'Unlock {path.name} before merging.')
                    merged.insert_pdf(pdf)
            save_pdf(merged, target)
        return

    with fitz.open(str(sources[0])) as pdf:
        encrypted = pdf.needs_pass or bool(pdf.metadata.get('encryption'))
        if tool == 'Unlock PDF':
            if not encrypted:
                raise ValueError('This PDF is already unencrypted.')
            if not pdf.authenticate(option) & 4:
                raise ValueError('Enter the PDF owner password to unlock it.')
            save_pdf(pdf, target, encryption=fitz.PDF_ENCRYPT_NONE)
            return
        if encrypted:
            raise ValueError('Unlock this PDF before using this tool.')
        if tool == 'PDF to Excel':
            export_tables(pdf, target)
        elif tool in ('PDF to Text', 'PDF to Word'):
            export_text(pdf, target, tool == 'PDF to Word')
        elif tool == 'Split PDF':
            with zipfile.ZipFile(target, 'w') as archive:
                for i in range(pdf.page_count):
                    with fitz.open() as page:
                        page.insert_pdf(pdf, from_page=i, to_page=i)
                        archive.writestr(f'page-{i+1:05d}.pdf', page.tobytes(garbage=4, deflate=True))
        elif tool == 'Protect PDF':
            if not option or len(option.encode('utf-8')) > 40:
                raise ValueError('Enter a password between 1 and 40 UTF-8 bytes.')
            save_pdf(pdf, target, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=option,
                     user_pw=option, permissions=-1)
        else:
            if tool == 'Rotate PDF':
                for page in pdf:
                    page.set_rotation((page.rotation + int(option)) % 360)
            elif tool == 'Extract / reorder pages':
                pdf.select(page_order(option, pdf.page_count))
            elif tool == 'Remove pages':
                removed = set(page_order(option, pdf.page_count))
                remaining = [index for index in range(pdf.page_count) if index not in removed]
                if not remaining:
                    raise ValueError('Keep at least one page.')
                pdf.select(remaining)
            save_pdf(pdf, target)
            if tool == 'Compress PDF' and target.stat().st_size >= sources[0].stat().st_size:
                raise ValueError('Lossless optimization could not make this PDF smaller.')


def convert(tool, sources, destination, option):
    with tempfile.TemporaryDirectory(prefix='.localpdf-', dir=destination.parent) as folder:
        staged = Path(folder) / ('converted' + TOOLS[tool][1])
        try:
            if tool == 'Images to PDF':
                images_to_pdf(sources, staged)
            elif tool == 'PDF to JPG / PNG':
                pdf_to_images(sources[0], staged, option)
            elif tool == 'PDF to PowerPoint':
                pdf_to_slides(sources[0], staged)
            elif tool in OFFICE:
                check_office_file(sources[0])
                run_script('Office.ps1', '-Source', sources[0], '-Destination', staged, '-Mode', OFFICE[tool])
            else:
                edit_pdf(tool, sources, staged, option)
        except ModuleNotFoundError as error:
            raise ValueError('Run Install-Extras.cmd once, then restart LocalPDF.') from error
        if not staged.is_file() or not staged.stat().st_size:
            raise ValueError('The converter did not create an output file.')
        # Windows rename is atomic on the same volume and fails if the destination exists.
        staged.rename(destination)



class LocalPDF:
    def __init__(self, root):
        self.root = root
        self.paths = []
        self.events = queue.Queue()
        self.busy = False
        self.tool = tk.StringVar(value=next(iter(TOOLS)))
        self.option = tk.StringVar()
        self.status = tk.StringVar()
        self.output = None

        root.title('LocalPDF')
        root.geometry('965x690')
        root.minsize(820, 620)
        root.configure(bg='#edf3f6')
        style = ttk.Style(root)
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10), padding=(11, 7))
        style.configure('TCombobox', padding=6)
        style.configure('Convert.TButton', background='#086768', foreground='white', font=('Segoe UI Semibold', 11))
        style.map('Convert.TButton', background=[('disabled', '#b8cdcf'), ('active', '#075657')])

        header = tk.Frame(root, bg='#142c39', padx=27, pady=19)
        header.pack(fill='x')
        tk.Label(header, text='LocalPDF', bg='#142c39', fg='white',
                 font=('Segoe UI Semibold', 26)).pack(side='left')
        tk.Label(header, text='Offline file conversion', bg='#142c39', fg='#b4dce0',
                 font=('Segoe UI', 11)).pack(side='right')

        body = tk.Frame(root, bg='#edf3f6', padx=20, pady=20)
        body.pack(fill='both', expand=True)
        rail = tk.Frame(body, bg='#e0ebef', padx=12, pady=12)
        rail.pack(side='left', fill='y', padx=(0, 16))
        tk.Label(rail, text='Tools', bg='#e0ebef', fg='#142c39',
                 font=('Segoe UI Semibold', 13)).pack(anchor='w', pady=(0, 7))
        self.menu = tk.Listbox(rail, width=25, height=len(TOOLS), selectmode='browse', exportselection=False,
                               borderwidth=0, highlightthickness=0, bg='#e0ebef', fg='#142c39',
                               selectbackground='#086768', selectforeground='white',
                               font=('Segoe UI', 11), activestyle='none')
        self.menu.pack(fill='both', expand=True)
        for title in TOOLS:
            self.menu.insert('end', title)
        self.menu.selection_set(0)
        self.menu.bind('<<ListboxSelect>>', self.change_tool)

        card = tk.Frame(body, bg='white', padx=24, pady=22)
        card.pack(side='left', fill='both', expand=True)
        tk.Label(card, textvariable=self.tool, bg='white', fg='#142c39',
                 font=('Segoe UI Semibold', 22)).pack(anchor='w')
        self.description = tk.Label(card, bg='white', fg='#435b67', justify='left', anchor='w',
                                    font=('Segoe UI', 11), wraplength=490)
        self.description.pack(fill='x', pady=(6, 18))
        self.description.bind('<Configure>', lambda event: self.description.configure(wraplength=event.width))
        controls = tk.Frame(card, bg='white')
        controls.pack(fill='x', pady=(0, 8))
        self.add = ttk.Button(controls, text='Add files', command=self.add_files)
        self.add.pack(side='left')
        self.order_buttons = []
        for title, action in [('Remove', lambda: self.move(0)), ('Up', lambda: self.move(-1)), ('Down', lambda: self.move(1))]:
            button = ttk.Button(controls, text=title, command=action)
            button.pack(side='left', padx=(5, 0))
            self.order_buttons.append(button)
        file_box = tk.Frame(card, bg='#f2f7f8')
        file_box.pack(fill='both', expand=True)
        self.files = tk.Listbox(file_box, height=5, bg='#f2f7f8', fg='#142c39', relief='flat',
                                highlightthickness=0, font=('Segoe UI', 10), exportselection=False)
        self.files.pack(side='left', fill='both', expand=True, padx=(8, 0), pady=5)
        scroll = ttk.Scrollbar(file_box, orient='vertical', command=self.files.yview)
        scroll.pack(side='right', fill='y')
        self.files.configure(yscrollcommand=scroll.set)
        self.option_area = tk.Frame(card, bg='white')
        self.option_area.pack(fill='x')
        self.option_label = tk.Label(self.option_area, bg='white', fg='#142c39', font=('Segoe UI Semibold', 10))
        self.option_entry = ttk.Entry(self.option_area, textvariable=self.option)
        self.option_menu = ttk.Combobox(self.option_area, textvariable=self.option, state='readonly')
        footer = tk.Frame(card, bg='white')
        footer.pack(fill='x', pady=(16, 0))
        self.run = ttk.Button(footer, text='Convert and save', style='Convert.TButton', command=self.start)
        self.run.pack(side='left')
        self.open = ttk.Button(footer, text='Open output folder', command=self.open_output, state='disabled')
        self.open.pack(side='left', padx=7)
        self.status_label = tk.Label(card, textvariable=self.status, bg='white', fg='#435b67',
                                     justify='left', anchor='w', wraplength=480, font=('Segoe UI', 10))
        self.status_label.pack(fill='x', pady=(10, 0))
        self.status_label.bind('<Configure>', lambda event: self.status_label.configure(wraplength=event.width))
        self.change_tool()
        root.protocol('WM_DELETE_WINDOW', self.close)
        root.after(100, self.poll)

    def change_tool(self, _event=None):
        selection = self.menu.curselection()
        if not selection or self.busy:
            return
        tool = self.menu.get(selection[0])
        patterns, _, multiple, label, description = TOOLS[tool]
        self.tool.set(tool)
        self.description.configure(text=description)
        self.paths = [path for path in self.paths if '*' + path.suffix.lower() in patterns]
        if not multiple:
            self.paths = self.paths[:1]
        self.repaint_files()
        self.add.configure(text='Add files' if multiple else 'Choose file')
        for widget in self.option_area.winfo_children():
            widget.pack_forget()
        self.option.set('')
        if label:
            self.option_label.configure(text=label)
            self.option_label.pack(anchor='w', pady=(10, 2))
            if tool in CHOICES:
                self.option_menu.configure(values=CHOICES[tool])
                self.option_menu.current(0)
                self.option_menu.pack(anchor='w')
            else:
                self.option_entry.configure(show='*' if 'password' in label.lower() else '')
                self.option_entry.pack(fill='x')
        self.status.set('Choose source files to begin.')

    def repaint_files(self):
        self.files.delete(0, 'end')
        for path in self.paths:
            self.files.insert('end', str(path))

    def add_files(self):
        patterns, _, multiple, *_ = TOOLS[self.tool.get()]
        kinds = [('Supported files', ' '.join(patterns))]
        picked = filedialog.askopenfilenames(filetypes=kinds) if multiple else filedialog.askopenfilename(filetypes=kinds)
        if not picked:
            return
        paths = [Path(name).resolve() for name in (picked if multiple else (picked,))]
        if any('*' + path.suffix.lower() not in patterns for path in paths):
            messagebox.showerror('Unsupported input', 'Choose a file with a supported extension.')
            return
        if multiple:
            self.paths.extend(path for path in paths if path not in self.paths)
        else:
            self.paths = paths
        self.repaint_files()
        self.status.set(f'{len(self.paths)} file(s) selected.')

    def move(self, offset):
        selection = self.files.curselection()
        if not selection:
            return
        index = selection[0]
        if not offset:
            self.paths.pop(index)
        elif 0 <= index + offset < len(self.paths):
            other = index + offset
            self.paths[index], self.paths[other] = self.paths[other], self.paths[index]
            index = other
        self.repaint_files()
        self.files.selection_set(min(index, len(self.paths) - 1))

    def start(self):
        tool = self.tool.get()
        if not self.paths or (tool == 'Merge PDFs' and len(self.paths) < 2):
            self.status.set('Choose at least two PDFs to merge.' if tool == 'Merge PDFs' else 'Choose a source file first.')
            return
        if tool not in FIXED and importlib.util.find_spec('pymupdf') is None:
            messagebox.showinfo('Setup needed', 'Run Install-Extras.cmd once to enable this tool.')
            return
        extension = TOOLS[tool][1]
        source = self.paths[0]
        name = filedialog.asksaveasfilename(initialdir=source.parent, initialfile=source.stem + '-converted' + extension,
                                           defaultextension=extension, filetypes=[(extension[1:].upper(), '*' + extension)],
                                           confirmoverwrite=False)
        if not name:
            return
        destination = Path(name).resolve()
        if destination.suffix.lower() != extension:
            messagebox.showerror('Output filename', f'Use the {extension} extension.')
            return
        if destination.exists():
            messagebox.showerror('Output filename', 'That filename already exists. Choose a new name.')
            return
        option = self.option.get()
        if tool in ('Protect PDF', 'Unlock PDF'):
            self.option.set('')
        self.set_busy(True)
        self.status.set('Converting on this PC…')
        threading.Thread(target=self.worker, args=(tool, tuple(self.paths), destination, option), daemon=True).start()

    def worker(self, tool, sources, destination, option):
        try:
            convert(tool, sources, destination, option)
        except Exception as error:
            self.events.put((None, str(error)))
        else:
            self.events.put((destination, ''))

    def set_busy(self, busy):
        self.busy = busy
        state = 'disabled' if busy else 'normal'
        for widget in (self.menu, self.add, *self.order_buttons, self.run, self.option_entry):
            widget.configure(state=state)
        self.option_menu.configure(state='disabled' if busy else 'readonly')

    def poll(self):
        try:
            output, error = self.events.get_nowait()
        except queue.Empty:
            pass
        else:
            self.set_busy(False)
            if output:
                self.output = output
                self.open.configure(state='normal')
                self.status.set(f'Saved {output.name}')
            else:
                self.status.set('Conversion failed.')
                messagebox.showerror('Conversion failed', error)
        self.root.after(100, self.poll)

    def open_output(self):
        os.startfile(self.output.parent)

    def close(self):
        if self.busy:
            self.status.set('Wait for the current conversion to finish before closing.')
        else:
            self.root.destroy()


if __name__ == '__main__':
    window = tk.Tk()
    LocalPDF(window)
    window.mainloop()


