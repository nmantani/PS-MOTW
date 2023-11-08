<#
Get-MOTW.ps1: PowerShell script to show MOTW (Mark of the Web)

Copyright (c) 2022, Nobutaka Mantani
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
   this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS
IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>

<#
.SYNOPSIS
Shows MOTW (Mark of the Web).

.DESCRIPTION
Get-MOTW.ps1 shows MOTW (Mark of the Web) of specified files. If a directory is specified, all files under the directory are processed recursively. The * wildcard can be used to specify multiple files. Only the "-Verbose" parameter is supported in CommonParameters.

.PARAMETER Path
Specifies the path to show MOTW. This parameter is mandatory.

.EXAMPLE
.\Get-MOTW.ps1 example.docx
C:\Users\user\Desktop\example.docx:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/

Description
---------------------------------------------------------------
Showing MOTW of example.docx.

.EXAMPLE
.\Get-MOTW.ps1 *.docx
C:\Users\user\Desktop\example1.docx:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/

C:\Users\user\Desktop\example2.docx:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/

Description
---------------------------------------------------------------
Showing MOTW of multiple Word document files.

.EXAMPLE
.\Get-MOTW.ps1 C:\Users\user\Documents
C:\Users\user\Documents\word\example.docx:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/

C:\Users\user\Documents\excel\example.xlsx:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/

Description
---------------------------------------------------------------
Showing MOTW of all files under C:\Users\user\Documents .
#>

Param(
        [parameter(mandatory=$true, ValueFromRemainingArguments=$true)]$Paths
)

foreach ($p in $Paths) {
    if (!(Test-Path $p)) {
        Write-Output "Error: $p does not exist."
        exit
    } elseif (Test-Path $p -PathType Container) {
        $files += @(Get-ChildItem -Force -Path $p -Recurse | Select-Object -ExpandProperty FullName)
    } else {
        $files += @(Resolve-Path $p)
    }
}

$count = 0
foreach ($f in $files) {
    $have_motw = $false
    $streams = Get-Item -Force -Stream * $f | Select-Object Stream
    foreach ($s in $streams) {
        if ($s.Stream -eq "Zone.Identifier") {
            $have_motw = $true
        }
    }

    if ($have_motw) {
        Write-Output "${f}:"
        Get-Content -Path $f -Stream Zone.Identifier -Encoding oem

        if ($count -lt ($files.Length - 1)) {
            Write-Output ""
        }
    } elseif ($VerbosePreference -eq "Continue") {
        Write-Output "$f does not have MOTW (Mark of the Web)."
    }

    $count += 1
}
