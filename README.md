# LocalPDF for Windows

LocalPDF is a desktop converter for Windows 10 and 11. It runs in a native window and processes files on your PC. There is no website, account, local web server, or upload step.

## Start

1. Extract the ZIP to a local folder.
2. Double-click **Start-LocalPDF.cmd**. Python 3.10 or later with Tcl/Tk is required.
3. Select a tool on the left, choose your input files, and click **Convert and save**.
4. Choose a new filename. The app places the result in your chosen folder and keeps the originals intact.

For merge and image-to-PDF jobs, use **Up** and **Down** to set the desired order. Use **Open output folder** to locate a completed result.

## Tools and requirements

| Tools | Requirements |
| --- | --- |
| JPG, PNG, WebP, BMP, TIFF to PDF | Pillow |
| PDF to JPG or PNG | Pillow and the Windows PDF renderer. Produces a ZIP with one image per page. Very large pages are capped at 6,000 pixels on the long side. |
| Word, Excel or PowerPoint to PDF | The matching, activated desktop Microsoft Office app. Accepts DOCX, XLSX and PPTX. |
| PDF to Word or plain text | PyMuPDF. Word export also uses python-docx. Extracts selectable text. Original layout and images are omitted. |
| PDF to Excel | PyMuPDF and openpyxl. Exports detected tables as text cells. Works best with visible table borders. |
| PDF to PowerPoint | Pillow and python-pptx. Produces one image per slide. |
| Merge, split, extract, reorder, remove, rotate, compress, protect, unlock | PyMuPDF. Split produces a ZIP. Compress uses lossless optimization. Unlock requires the owner password. |

Double-click **Install-Extras.cmd** once if a tool asks for setup. It uses pip to download the Python packages. If your Python version cannot install the packages, use a supported 64-bit Python release such as Python 3.13 and run the launcher with that same interpreter. Conversion can then run offline.

## Handling confidential files

LocalPDF has no upload or telemetry code. Save to a regular local folder to avoid OneDrive or other folder sync. For Office conversion, you can disconnect the PC from the network after setup to prevent network activity from Office or other software. Office conversion disables macros and screens the document package for external content links.

PDF-to-Word and PDF-to-Excel require a PDF with a text layer. Scans without selectable text need OCR separately. PDF editing can affect forms, bookmarks, and digital signatures. Keep the original if you need those features.

## Verification

The desktop interface and the image/PDF conversions were tested on Windows with synthetic files. The test checked two-page order, PNG and JPEG output, image size, transparency against white, and existing-file protection. Package downloads and Office COM activation were blocked in the testing environment, so the optional PDF tools and Office conversions could not be verified end to end.
