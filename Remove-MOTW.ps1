<#
Remove-MOTW.ps1: PowerShell script to remove MOTW (Mark of the Web)

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
Removes MOTW (Mark of the Web).

.DESCRIPTION
Remove-MOTW.ps1 removes MOTW (Mark of the Web) from speficied files. If a directory is specified, all files under the directory are processed recursively. The * wildcard can be used to specify multiple files. Only the "-Verbose" parameter is supported in CommonParameters.

.PARAMETER Path
Specifies the path to remove MOTW. This parameter is mandatory.

.EXAMPLE
.\Remove-MOTW.ps1 example.docx

Description
---------------------------------------------------------------
Removing MOTW from example.docx.

.EXAMPLE
.\Remove-MOTW.ps1 *.jpg

Description
---------------------------------------------------------------
Removing MOTW from multiple JPEG files.

.EXAMPLE
.\Remove-MOTW.ps1 C:\Users\user\Downloads

Description
---------------------------------------------------------------
Removing MOTW from all files under C:\Users\user\Downloads .
#>

Param(
    [parameter(mandatory=$true)][String]$Path
)

function dummy {
    
}

if (!(Test-Path $Path)) {
    Write-Host "Error: $Path does not exist."
    exit
} elseif (Test-Path $Path -PathType Container) {
    $files = Get-ChildItem -Path $Path -Recurse | Select-Object -ExpandProperty FullName
} else {
    $files = Resolve-Path $Path
}

foreach ($f in $files) {
    $have_motw = $false
    $streams = Get-Item -Stream * $f | Select-Object Stream
    foreach ($s in $streams) {
        if ($s.Stream -eq "Zone.Identifier") {
            $have_motw = $true
        }
    }

    if ($have_motw) {
        Remove-Item -Stream Zone.Identifier $f
    } elseif ($VerbosePreference -eq "Continue") {
        Write-Host "${f} does not have MOTW (Mark of the Web)."
    }
}

