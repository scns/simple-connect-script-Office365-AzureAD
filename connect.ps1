<#
#### requires ps-version 5.1 ####

.DESCRIPTION
Default connection script for O365

.NOTES
   Version:        0.1
   Author:         Maarten Schmeitz
   Creation Date:  Wednesday, November 6th 2019, 10:20:03 am
   File: General_Connect_Office365.ps1
   Copyright (c) 2019 Advantive

HISTORY:
Date      	          By	Comments
----------	          ---	----------------------------------------------------------

.LINK
   www.advantive.nl

.LICENSE
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the Software), to deal
in the Software without restriction, including without limitation the rights
to use copy, modify, merge, publish, distribute sublicense and /or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
#>


if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host "MSOnline Module exists"
} 
else {
   install-module MSOnline
}

if (Get-Module -ListAvailable -Name AzureAD) {
    Write-Host "AzureAD Module exists"
} 
else {
   install-module azureAD
}

if (!(Get-PSSession | Where-object { $_.ConfigurationName -eq "Microsoft.Exchange" })) { 
    Set-ExecutionPolicy RemoteSigned
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    $host.ui.rawui.windowtitle = $UserCredential.UserName
    Import-PSSession $Session
    Connect-MsolService -Credential $UserCredential
    Connect-AzureAD -Credential $UserCredential

}
