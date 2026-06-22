# TeklifAgent COM DLL derleme (.NET Framework 4.x csc)
# Çıktı: data/agent/TeklifAgent.Com.x64.dll ve TeklifAgent.Com.x86.dll

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Src  = Join-Path $PSScriptRoot "src"
$Out  = Join-Path $Root "data\agent"
New-Item -ItemType Directory -Force -Path $Out | Out-Null

$Files = @(
    (Join-Path $Src "AgentConfig.cs"),
    (Join-Path $Src "ApiClient.cs"),
    (Join-Path $Src "ExcelRunner.cs"),
    (Join-Path $Src "AgentWorker.cs"),
    (Join-Path $Src "AgentLog.cs"),
    (Join-Path $Src "AgentControl.cs"),
    (Join-Path $Src "Program.cs")
)

$Ref = "/reference:System.Web.Extensions.dll"

function Build-Agent($CscPath, $Platform, $OutName) {
    Write-Host "Derleniyor: $OutName ($Platform) ..."
    $outFile = Join-Path $Out $OutName
    $args = @(
        "/target:library",
        "/platform:$Platform",
        "/out:$outFile",
        $Ref
    ) + $Files
    & $CscPath @args
    if ($LASTEXITCODE -ne 0) { throw "csc hatasi: $OutName" }
}

$Csc64 = "${env:WINDIR}\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$Csc86 = "${env:WINDIR}\Microsoft.NET\Framework\v4.0.30319\csc.exe"

Build-Agent $Csc64 "anycpu" "TeklifAgent.Com.x64.dll"
Copy-Item (Join-Path $Out "TeklifAgent.Com.x64.dll") (Join-Path $Out "TeklifAgent.Com.dll") -Force

if (Test-Path $Csc86) {
    Build-Agent $Csc86 "x86" "TeklifAgent.Com.x86.dll"
}

# Exe (worker fallback)
$exeOut = Join-Path $Out "TeklifAgent.exe"
& $Csc64 /target:exe /platform:anycpu "/out:$exeOut" $Ref @Files
if ($LASTEXITCODE -ne 0) { throw "csc exe hatasi" }

Write-Host ""
Write-Host "Tamamlandi: $Out"
Write-Host "  TeklifAgent.Com.dll      (x64 Office icin)"
Write-Host "  TeklifAgent.Com.x86.dll  (32-bit Office icin)"
Write-Host "  TeklifAgent.exe          (--worker modu)"
Write-Host ""
Write-Host "COM kayit: regasm `"$Out\TeklifAgent.Com.dll`" /codebase"
