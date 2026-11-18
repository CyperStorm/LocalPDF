param([string]$Source,[string]$Destination,[string]$Mode)
$ErrorActionPreference = 'Stop'
$app = $null
$doc = $null
try {
    switch ($Mode) {
        'word' {
            $app = New-Object -ComObject Word.Application
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 0
            $app.Options.UpdateLinksAtOpen = $false
            $app.Options.UpdateLinksAtPrint = $false
            $app.Options.UpdateFieldsAtPrint = $false
            $doc = $app.Documents.OpenNoRepairDialog($Source, $false, $true, $false)
            $doc.ExportAsFixedFormat($Destination, 17)
        }
        'excel' {
            $app = New-Object -ComObject Excel.Application
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = $false
            $app.EnableEvents = $false
            $app.AskToUpdateLinks = $false
            $doc = $app.Workbooks.Open($Source, 0, $true)
            $doc.ExportAsFixedFormat(0, $Destination)
        }
        'powerpoint' {
            $app = New-Object -ComObject PowerPoint.Application
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 1
            $doc = $app.Presentations.Open($Source, -1, 0, 0)
            $doc.SaveAs($Destination, 32)
        }
        default { throw 'Unknown Office conversion.' }
    }
} finally {
    if ($doc) {
        if ($Mode -eq 'powerpoint') { $doc.Close() } else { $doc.Close(0) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($app) {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}
