$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& "$PSScriptRoot\make-double-w.ps1"
& $compiler /nologo /target:winexe /optimize+ /win32icon:"$PSScriptRoot\assets\double-w.ico" /resource:"$PSScriptRoot\assets\double-w.ico",AppIcon /out:"$PSScriptRoot\Word文本比对器新版.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:Microsoft.CSharp.dll "$PSScriptRoot\WordCompare.cs"
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
