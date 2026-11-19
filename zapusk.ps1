$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$app = Join-Path $here "windows-pult"
if (-not (Test-Path (Join-Path $app "pult_gui.pyw"))) { $app = $here }
$log = Join-Path $here "zapusk.log"
function Note($text) {
    $line = "{0} {1}" -f (Get-Date -Format "HH:mm:ss"), $text
    Add-Content -Path $log -Value $line -Encoding UTF8
    Write-Host $line
}
Set-Content -Path $log -Value "=== zapusk $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -Encoding UTF8

function Test-Python($exe) {
    if (-not $exe) { return $false }
    if ($exe -match "\\WindowsApps\\") { return $false }
    if (-not (Test-Path -LiteralPath $exe)) { return $false }
    try {
        $out = & $exe -c "import sys; print(sys.executable)" 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and $out -match "[A-Z]:\\") {
            Note "rabochij: $($out.Trim())"
            return $true
        }
    } catch {}
    return $false
}

function Find-Python {
    $list = New-Object System.Collections.Generic.List[string]
    $roots = @(
        "$env:LOCALAPPDATA\Programs\Python",
        "$env:LOCALAPPDATA\Python",
        "$env:ProgramFiles",
        "${env:ProgramFiles(x86)}"
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        Get-ChildItem -Path $root -Filter python.exe -Recurse -Depth 5 -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -notmatch "WindowsApps" } |
            ForEach-Object { $list.Add($_.FullName) }
    }
    foreach ($c in ($list | Select-Object -Unique)) {
        if (Test-Python $c) { return $c }
    }
    return $null
}

$pyexe = Find-Python
if (-not $pyexe) {
    Note "nastoyaschego python.exe net. stavlju cherez winget"
    Write-Host "Stavlju Python. Eto zanimaet neskolko minut, ne zakryvaj okno."
    & winget install --id Python.Python.3.14 -e --accept-package-agreements --accept-source-agreements --disable-interactivity
    Note "winget zakonchil, kod $LASTEXITCODE"
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    $pyexe = Find-Python
}
if (-not $pyexe) {
    Note "probuju py install 3.14"
    & py install 3.14
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    $pyexe = Find-Python
}
if (-not $pyexe) {
    Note "Python tak i ne vstal"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "С сайта скачался установщик, но сам Python не установился.`nФайла python.exe на диске нет — поэтому пульт и не открывается.`n`nЗапусти ЗАПУСК ПУЛЬТА.bat ещё раз и не закрывай чёрное окно, пока оно пишет Stavlju Python.",
        "Пульт",
        "OK",
        "Error"
    ) | Out-Null
    exit 1
}

Note "stavlju biblioteki"
& $pyexe -m pip install -r (Join-Path $app "requirements.txt")
if ($LASTEXITCODE -ne 0) { Note "pip ne udalsja"; exit 1 }
Note "playwright chromium"
& $pyexe -m playwright install chromium
if ($LASTEXITCODE -ne 0) { Note "playwright ne udalsja"; exit 1 }

$dir = Split-Path $pyexe
$pyw = Join-Path $dir "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $pyexe }
Note "okno: $pyw"
$desk = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut((Join-Path $desk "Пульт.lnk"))
$link.TargetPath = $pyw
$link.Arguments = '"' + (Join-Path $app "pult_gui.pyw") + '"'
$link.WorkingDirectory = $app
$ico = Join-Path $app "pult.ico"
if (Test-Path $ico) { $link.IconLocation = $ico }
$link.Save()
Start-Process -FilePath $pyw -ArgumentList ('"' + (Join-Path $app "pult_gui.pyw") + '"') -WorkingDirectory $app
Note "zapustil"
exit 0
