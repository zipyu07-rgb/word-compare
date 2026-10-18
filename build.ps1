$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /optimize+ /out:"$PSScriptRoot\WordCompare.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:Microsoft.CSharp.dll "$PSScriptRoot\WordCompare.cs"
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
