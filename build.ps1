<#
.SYNOPSIS
    Build the StockData LibreOffice Calc add-in (.oxt).

.DESCRIPTION
    1. Compiles src/idl/XStockData.idl into a UNO type library (StockData.rdb)
       using the LibreOffice SDK's unoidl-write (modern) or idlc + regmerge
       (legacy) tool.
    2. Stages the .oxt payload (description.xml, manifest, python component,
       config and the generated .rdb).
    3. Zips the staging folder into dist/StockData.oxt.

    Requires the LibreOffice SDK. Point -SdkDir / -OfficeDir at your install,
    or set the LO_SDK_HOME / LO_PROGRAM environment variables.

.EXAMPLE
    .\build.ps1 -OfficeDir "C:\Program Files\LibreOffice\program" `
                -SdkDir    "C:\Program Files\LibreOffice\sdk"
#>
[CmdletBinding()]
param(
    [string]$OfficeDir = $env:LO_PROGRAM,
    [string]$SdkDir    = $env:LO_SDK_HOME
)

$ErrorActionPreference = "Stop"
$root  = $PSScriptRoot
$build = Join-Path $root "build"
$stage = Join-Path $build "oxt"
$dist  = Join-Path $root "dist"

function Find-FirstExisting([string[]]$candidates) {
    foreach ($c in $candidates) { if ($c -and (Test-Path $c)) { return $c } }
    return $null
}

# --- locate the toolchain ----------------------------------------------------
if (-not $OfficeDir) {
    $OfficeDir = Find-FirstExisting @(
        "C:\Program Files\LibreOffice\program",
        "C:\Program Files (x86)\LibreOffice\program"
    )
}
if (-not $SdkDir) {
    $SdkDir = Find-FirstExisting @(
        "C:\Program Files\LibreOffice\sdk",
        "C:\Program Files (x86)\LibreOffice\sdk"
    )
}
if (-not $OfficeDir) { throw "LibreOffice program dir not found. Pass -OfficeDir." }

# The SDK tools (unoidl-write/idlc) link against the URE/UNO DLLs that live in
# the office "program" directory, so it must be on PATH or they fail to start
# with 0xC0000135 (DLL not found).
$env:PATH = "$OfficeDir;$env:PATH"

$typesRdb = Find-FirstExisting @(
    (Join-Path $OfficeDir "types.rdb"),
    (Join-Path $OfficeDir "types\offapi.rdb")
)
if (-not $typesRdb) { throw "Could not find types.rdb under $OfficeDir." }

$unoidlWrite = Find-FirstExisting @(
    (Join-Path $OfficeDir "unoidl-write.exe"),
    (Join-Path $SdkDir   "bin\unoidl-write.exe")
)
$idlc = if ($SdkDir) { Find-FirstExisting @((Join-Path $SdkDir "bin\idlc.exe")) } else { $null }

# --- clean staging -----------------------------------------------------------
Remove-Item -Recurse -Force $build, $dist -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $stage, $dist | Out-Null

$idl    = Join-Path $root "src\idl\XStockData.idl"
$rdbOut = Join-Path $stage "StockData.rdb"

# --- compile the IDL ---------------------------------------------------------
if ($unoidlWrite) {
    Write-Host "Compiling IDL with unoidl-write..." -ForegroundColor Cyan
    & $unoidlWrite $typesRdb $idl $rdbOut
    if ($LASTEXITCODE -ne 0) { throw "unoidl-write failed ($LASTEXITCODE)." }
}
elseif ($idlc) {
    Write-Host "Compiling IDL with idlc + regmerge..." -ForegroundColor Cyan
    $regmerge = Join-Path $OfficeDir "regmerge.exe"
    if (-not (Test-Path $regmerge)) { $regmerge = Join-Path $SdkDir "bin\regmerge.exe" }
    $urd = Join-Path $build "XStockData.urd"
    & $idlc -w -I (Join-Path $SdkDir "idl") -O $build $idl
    if ($LASTEXITCODE -ne 0) { throw "idlc failed ($LASTEXITCODE)." }
    & $regmerge $rdbOut /UCR $urd
    if ($LASTEXITCODE -ne 0) { throw "regmerge failed ($LASTEXITCODE)." }
}
else {
    throw "Neither unoidl-write nor idlc found. Install the LibreOffice SDK."
}

# --- stage payload -----------------------------------------------------------
Write-Host "Staging .oxt payload..." -ForegroundColor Cyan
Copy-Item (Join-Path $root "oxt\description.xml")  $stage
Copy-Item (Join-Path $root "oxt\CalcAddIns.xcu")   $stage
Copy-Item (Join-Path $root "oxt\Jobs.xcu")         $stage
Copy-Item (Join-Path $root "src\python\stockdata.py") $stage
Copy-Item (Join-Path $root "oxt\META-INF") $stage -Recurse
Copy-Item (Join-Path $root "oxt\description") $stage -Recurse
Copy-Item (Join-Path $root "oxt\icons") $stage -Recurse
Copy-Item (Join-Path $root "oxt\registration") $stage -Recurse

# --- zip into .oxt -----------------------------------------------------------
$oxt = Join-Path $dist "StockData.oxt"
Write-Host "Packing $oxt..." -ForegroundColor Cyan
if (Test-Path $oxt) { Remove-Item $oxt }
# Compress the *contents* of the staging dir to the archive root.
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $oxt -Force

Write-Host "Done: $oxt" -ForegroundColor Green
Write-Host "Install with:  unopkg add --force `"$oxt`"" -ForegroundColor Green
