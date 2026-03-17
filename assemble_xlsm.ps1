param(
    [Parameter(Mandatory = $true)][string]$dllPath,
    [Parameter(Mandatory = $true)][string]$basDir,
    [Parameter(Mandatory = $true)][string]$outPath
)

$ErrorActionPreference = "Stop"
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$excel = $null
$workbook = $null
$tempFiles = @()

try {
    $distFolder = Split-Path -Parent $outPath
    if (-not (Test-Path $distFolder)) {
        New-Item -ItemType Directory -Path $distFolder -Force | Out-Null
    }

    Copy-Item -Path $dllPath -Destination (Join-Path $distFolder "excel_deno_bridge.dll") -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    $templates = @("DenoCore.bas.template", "UserScripts.bas.template")

    foreach ($tempName in $templates) {
        $templatePath = Join-Path $basDir $tempName
        if (-not (Test-Path $templatePath)) { throw "Template not found: $templatePath" }

        $content = Get-Content -Path $templatePath -Raw
        $content = $content.Replace("{{GEN_DATE}}", (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

        $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $tempName.Replace(".template", ""))
        [System.IO.File]::WriteAllText($tmpFile, $content, $Utf8NoBom)
        $tempFiles += $tmpFile

        $null = $workbook.VBProject.VBComponents.Import($tmpFile)
    }

    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $workbook.SaveAs($outPath, 52)
    Write-Host "Bundle Ready: $outPath"
}
catch {
    Write-Error "Build failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($workbook) { $workbook.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    foreach ($tmp in $tempFiles) {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}