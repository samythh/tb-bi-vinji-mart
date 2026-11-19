param([string]$file, [string]$pdf)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open($file, $false, $false)
    foreach ($t in $d.TablesOfContents) { $t.Update() }
    $d.Repaginate()
    $n = $d.ComputeStatistics(2)
    $r = $d.Content
    $f = $r.Find
    $null = $f.Execute('{PAGES}', $false, $false, $false, $false, $false, $true, 1, $false, [string]$n, 2)
    $d.Repaginate()
    foreach ($t in $d.TablesOfContents) { $t.Update() }
    $n2 = $d.ComputeStatistics(2)
    Write-Output "PAGES=$n -> $n2"
    $d.Save()
    $d.ExportAsFixedFormat($pdf, 17)
    $d.Close(0)
} finally { $w.Quit() }
