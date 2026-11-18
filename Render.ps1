param([string]$Source,[string]$Destination,[int]$Dpi=150)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
$operation = $methods | Where-Object { $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }
$action = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' }
function Await-Result($pending, [type]$type) {
    $task = $operation.MakeGenericMethod($type).Invoke($null, @($pending))
    $task.GetAwaiter().GetResult()
}
$file = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream([IO.File]::OpenRead($Source))
$pdf = Await-Result ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($file)) ([Windows.Data.Pdf.PdfDocument])
for ($i=0; $i -lt $pdf.PageCount; $i++) {
    $page = $pdf.GetPage($i)
    $memory = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    try {
        $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $scale = [Math]::Min($Dpi / 96.0, 6000.0 / [Math]::Max($page.Size.Width, $page.Size.Height))
        $options.DestinationWidth = [uint32][Math]::Max(1, $page.Size.Width * $scale)
        $options.DestinationHeight = [uint32][Math]::Max(1, $page.Size.Height * $scale)
        $pending = $page.RenderToStreamAsync($memory, $options)
        $null = $action.Invoke($null, @($pending)).GetAwaiter().GetResult()
        $memory.Seek(0)
        $name = Join-Path $Destination ('page-{0:D5}.png' -f ($i+1))
        $out = [IO.File]::Create($name)
        try { [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($memory).CopyTo($out) }
        finally { $out.Dispose() }
        [IO.File]::WriteAllText("$name.size", "$($options.DestinationWidth) $($options.DestinationHeight)")
    } finally { $memory.Dispose(); $page.Dispose() }
}



