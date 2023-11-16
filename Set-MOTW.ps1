<#
Set-MOTW.ps1: PowerShell script to set MOTW (Mark of the Web)

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
Sets MOTW (Mark of the Web).

.DESCRIPTION
Set-MOTW.ps1 sets MOTW for specified files. If a directory is specified, all files under the directory are processed recursively. The * wildcard can be used to specify multiple files. Only the "-Verbose" parameter is supported in CommonParameters.

.PARAMETER Path
Specifies the path to set MOTW. This parameter is mandatory. The "-Path" string can be omitted. Multiple paths can be specified with a comma-separated list.

.PARAMETER ZoneId
Specifies the ZoneId value (default: 3):
0: Local machine (URLZONE_LOCAL_MACHINE)
1: Local intranet (URLZONE_INTRANET)
2: Trusted sites (URLZONE_TRUSTED)
3: Internet (URLZONE_INTERNET)
4: Untrusted sites (URLZONE_UNTRUSTED)
This parameter is always set unless AppZoneId is specified.

.PARAMETER ReferrerUrl
Specifies the string for ReferrerUrl value of MOTW (default: undefined). Google Chrome, Microsoft Edge (Blink-based), and Mozilla Firefox set this value.

.PARAMETER HostUrl
Specifies the string for the HostUrl value of MOTW (default = undefined). Google Chrome, Microsoft Edge (Blink-based), and Mozilla Firefox set this value.

.PARAMETER HostIpAddress
Specifies the string for HostIpAddress of MOTW (default: undefined). Legacy Microsoft Edge (EdgeHTML-based) sets this value.

.PARAMETER LastWriterPackageFamilyName
Specifies the string for LastWriterPackageFamilyName of MOTW (default: undefined). Legacy Microsoft Edge (EdgeHTML-based) sets this value.

.PARAMETER AppZoneId
Specifies AppZoneId of MOTW (default: undefined). AppDefinedZoneId and ZoneId cannot be used if this parameter is specified. Old versions of SmartScreen set "AppZoneId=4" and remove ZoneId for an executable file when execution permission is given by clicking the "Run anyway" button. Recent versions of SmartScreen seem to just remove Zone.Identifier alternate data stream instead of setting "AppZoneId=4".

.PARAMETER AppDefinedZoneId
Specifies AppDefinedZoneId of MOTW (default: undefined). The purpose of AppDefinedZoneId is unknown and it is only mentioned in the "Zone.Identifier alternate data stream format" section of the document of IZoneIdentifier2 interface (https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/mt243886(v=vs.85)#zoneidentifier-alternate-data-stream-format).

.EXAMPLE
.\Set-MOTW.ps1 example.docx -ReferrerUrl https://example.com/ -HostUrl https://example.com/download/example.docx -Verbose
New MOTW (Mark of the Web) of C:\Users\user\Desktop\example.zip:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/example.docx

Description
---------------------------------------------------------------
Marking a Word document file as downloaded with web browsers.
New MOTW information is shown with -Verbose option.

.EXAMPLE
.\Set-MOTW.ps1 example.zip -ReferrerUrl https://example.net/ -HostUrl https://example.net/example.zip -HostIpAddress 192.168.100.100 -Verbose
Current MOTW (Mark of the Web) of C:\Users\user\Desktop\example.zip:
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.com/
HostUrl=https://example.com/download/example.zip

New MOTW (Mark of the Web) of C:\Users\user\Desktop\example.zip:
[ZoneTransfer]
HostIpAddress=192.168.100.100
ZoneId=3
ReferrerUrl=https://example.net/
HostUrl=https://example.net/example.zip

Description
---------------------------------------------------------------
Overwriting existing MOTW of example.zip with new MOTW to simulate the behavior of Legacy Microsoft Edge (EdgeHTML-based) when a file is downloaded with the "Save target as" context menu and saved to non-default location.

.EXAMPLE
.\Set-MOTW.ps1 *.jpg,*.png -ZoneId 2 -ReferrerUrl https://example.com/ -HostUrl https://example.com/download/

Description
---------------------------------------------------------------
Marking JPEG files and PNG files as downloaded from trusted sites (ZoneId = 2) with web browsers.

.EXAMPLE
.\Set-MOTW.ps1 example\*.png -ReferrerUrl C:\Users\user\Desktop\example.zip

Description
---------------------------------------------------------------
Simulating the behavior of "Extract all" built-in function of Windows Explorer that sets ReferrerUrl for extracted files to the path of a ZIP archive file.

.EXAMPLE
.\Set-MOTW.ps1 example.exe -AppZoneId 4

Description
---------------------------------------------------------------
Simulating the behavior of old versions of SmartScreen that set AppZoneId=4 for an executable file.

.EXAMPLE
.\Set-MOTW.ps1 C:\Users\user\Downloads -LastWriterPackageFamilyName Microsoft.Office.OneNote_8wekyb3d8bbwe -AppDefinedZoneId 0

Description
---------------------------------------------------------------
Marking all files under C:\Users\user\Downloads with the parameters LastWriterPackageFamilyName and AppDefinedZoneId.
#>

[CmdletBinding(PositionalBinding=$false)]

Param(
    [Int16]$ZoneId = -1,
    [String]$ReferrerUrl,
    [String]$HostUrl,
    [String]$HostIpAddress,
    [String]$LastWriterPackageFamilyName,
    [Int16]$AppZoneId = -1,
    [Int16]$AppDefinedZoneId = -1,
    [parameter(mandatory=$true, ValueFromRemainingArguments=$true)]$Path
)

if ($AppZoneId -ne -1) {
    if ($ZoneId -ne -1) {
        Write-Output "Error: AppZoneId and ZoneId cannot be used at the same time."
        exit
    } elseif ($AppDefinedZoneId -ne -1) {
        Write-Output "Error: AppDefinedZoneId and AppZoneId cannot be used at time same time."
        exit
    }
} elseif ($ZoneId -eq -1) {
    $ZoneId = 3
}

if ($ZoneId -ne -1 -and ($ZoneId -lt -1 -or $ZoneId -gt 4)) {
    Write-Output "Error: ZoneId is invalid."
    Write-Output "ZoneId has to be one of the following (default = 3):"
    Write-Output "0: Local machine (URLZONE_LOCAL_MACHINE)"
    Write-Output "1: Local intranet (URLZONE_INTRANET)"
    Write-Output "2: Trusted sites (URLZONE_TRUSTED)"
    Write-Output "3: Internet (URLZONE_INTERNET)"
    Write-Output "4: Untrusted sites (URLZONE_UNTRUSTED)"
    exit
}

foreach ($p in $Path) {
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

    if ($have_motw -and $VerbosePreference -eq "Continue") {
        Write-Output "Current MOTW (Mark of the Web) of ${f}:"
        Get-Content -Path $f -Stream Zone.Identifier -Encoding oem
        Write-Output ""
    }

    $motw = "[ZoneTransfer]`r`n"

    if ($HostIpAddress -ne "") {
        $motw += "HostIpAddress=$HostIpAddress`r`n"
    }

    if ($AppZoneId -ne -1) {
        $motw += "AppZoneId=$AppZoneId`r`n"
    }

    if ($ZoneId -ne -1) {
        $motw += "ZoneId=$ZoneId`r`n"
    }

    if ($LastWriterPackageFamilyName -ne "") {
        $motw += "LastWriterPackageFamilyName=$LastWriterPackageFamilyName`r`n"
    }

    if ($AppDefinedZoneId -ne -1) {
        $motw += "AppDefinedZoneId=$AppDefinedZoneId`r`n"
    }

    if ($ReferrerUrl -ne "") {
        $motw += "ReferrerUrl=$ReferrerUrl`r`n"
    }

    if ($HostUrl -ne "") {
        $motw += "HostUrl=$HostUrl`r`n"
    }

    if ($PSVersionTable.PSVersion.Major -lt 6 -and [Console]::OutputEncoding.CodePage -eq 65001) {
        # This is necessary to write Zone.Identfier without byte order mark on the environment with UTF-8 locale and PowerShell 5 or older
        $utf8nobom = New-Object System.Text.UTF8Encoding $false
        Set-Content -ErrorVariable error -Path $f -Stream Zone.Identifier -Encoding Byte -NoNewline -Value $utf8nobom.GetBytes($motw)
    } else {
        Set-Content -ErrorVariable error -Path $f -Stream Zone.Identifier -Encoding oem -NoNewline -Value $motw
    }

    if ($error -ne "") {
        $count += 1
        continue
    }

    if ($VerbosePreference -eq "Continue") {
        Write-Output "New MOTW (Mark of the Web) of ${f}:"
        Get-Content -Path $f -Stream Zone.Identifier -Encoding oem

        if ($count -lt ($files.Length - 1)) {
            Write-Output ""
        }
    }

    $count += 1
}
