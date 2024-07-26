<#
.SYNOPSIS
This script retrieves the bitlocker password for the specified computer, and renders the password as a barcode on screen. 

The barcode font is NOT INCLUDED, and will need to be seperately downloaded. A Code128 font will be required. This script was
developed with the Libre Barcode 128 font. 

https://fonts.google.com/specimen/Libre+Barcode+128

Download the font and place it in the same folder as this script. 


.NOTES
   File Name: bitlockerbarcode.ps1
   Author: Rob Woltz

.PARAMETER DomainController
   Optional. The script will use the default domain controller of the local machine. If you need to connect to another domain controller, specify here

.PARAMETER ComputerName.
   Optional. If a computer name is not specified when launching the script, the user will be prompted. 

#>

param (
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string] $DomainController,
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string] $ComputerName
)

#$fontName = "LibreBarcode128-Regular.ttf"

# Import the Barcode encoding script
import-module "$PSScriptRoot\code128\code128.psm1"

write-host "$PSScriptRoot\code128\code128.psm1"

# Prompt the user for a computername if not specified. 
if (!$ComputerName) {
    $computername = (read-host "Enter Hostname").trim()
}

$paramHash = @{
    Identity = $computername
    Properties = ('Name', 'DistinguishedName', 'ms-Mcs-AdmPwd')
 
}

if ($DomainController) {
    $paramHash.Add('Server',$DomainController)
}


# Get the computer object from Active Directory
try {
    $computer = get-adcomputer @paramHash
} catch {
    write-host "$computername not found in AD. Exiting..."
    read-host
    exit
}

[array]$ADBitlocker = @()

$paramHash = @{
    Filter = {objectclass -eq 'msFVE-RecoveryInformation'}
    Properties = 'msFVE-RecoveryPassword'
    SearchBase = $computer.distinguishedname
 
}

if ($DomainController) {
    $paramHash.Add('Server',$DomainController)
}

# Get the BitLocker Recovery keys from AD
foreach ($recoveryInfo in Get-ADObject @paramHash) {
    $recoveryKey = ($recoveryInfo.DistinguishedName | Select-String -Pattern 'CN=.+{(.+)},.+'  ).Matches.groups[1].value

    $Properties = @{'RecoveryKey'=$recoveryKey; 'RecoveryPassword'=$recoveryInfo.'msFVE-RecoveryPassword'}
    $ADBitlocker += New-Object -TypeName PSObject -Property $Properties
}

$selection = 0

# If multiple BitLocker keys are found, prompt the user to select the appropriate key

if ($ADBitlocker.Count -gt 1) {

    write-host "What are the first 8 digits of the bit locker recovery password?" 

    for ($i=0; $i -lt $ADBitlocker.count; $i++) {
        write-host "$($i): $($ADBitlocker[$i].RecoveryKey.SubString(0, 8))"
    }

    $selection = (Read-Host).Trim()

    try {
        [int]$selection | Out-Null
    } catch {
        write-host "Invalid Selection. Exiting..."
        Read-Host
        exit
    }
    if ([int]$selection -lt 0 -or [int]$selection -ge $i) {
        write-host "Invalid Selection. Exiting..."
        Read-Host
        exit
    }
}

$RPText = $ADBitlocker[$selection].RecoveryPassword

# Strip the dashes from the bitlocker password
$RP = $RPText.replace("-","")


$BarcodeArray = $null
[Array]$BarcodeArray = @()

# Convert the 48 digit bitlocker password to four seperate barcodes

0..3 | % {
    $BarcodeArray += Get-Code128String -Text $RP.Substring($_ * 12, 12) -Type B
}

$RecoveryBarcode = $BarcodeArray -join "`r`n`r`n"

$LAPS = $computer.'ms-Mcs-AdmPwd'

# Convert the LAPS password to a barcode. 

$LAPSEncode = Get-Code128String -Text $LAPS -Type B

#Generate a simple WPF form to display the barcodes

Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window"
    Width="600"
    Height="775"
    >
    <Grid x:Name="Grid">
        <Grid.Resources >
            <Style TargetType="Grid" >
                <Setter Property="Margin" Value="10" />
            </Style>
        </Grid.Resources>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Label x:Name = "BLInfo"
            Grid.Column="0"
            Grid.Row="0"
            Width="500"
            Content="BitLocker Code for $($computername.toupper())&#10;&#10;$RPText&#10;&#10;Scan the four barcodes below from top to bottom.&#10;&#10;If you're having issues scanning, the scanner might have trouble reading from your screen.&#10;Tilt the scanner, or try running scanning from an external monitor."
        />
        <Label x:Name = "BLBarcode"
            Grid.Column="0"
            Grid.Row="1"
            Width="500"
            FontSize="42"
            HorizontalContentAlignment="Center"
        />
        <Label x:Name = "LAPSInfo"
            Grid.Column="0"
            Grid.Row="2"
            Width="500"
            FontFamily="Consolas"
            FontSize="16"
        />
        <Label x:Name = "LAPSBarcode"
            Grid.Column="0"
            Grid.Row="3"
            Width="500"
            FontSize="42"
            HorizontalContentAlignment="Center"
        />
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)

$window = [Windows.Markup.XamlReader]::Load($reader)

#Set the values of the labels of the form
#And set the font to the Barcode font

$WPFBLBarcode = $window.FindName("BLBarcode")
$WPFBLBarcode.FontFamily = "$PSScriptRoot/LibreBarcode128-Regular.ttf#Libre Barcode 128"
$WPFBLBarcode.Content = $RecoveryBarcode

$WPFLAPSInfo = $window.FindName("LAPSInfo")
$WPFLAPSInfo.Content = "LAPS Password:`r`n$LAPS"

$WPFLAPSBarcode = $window.FindName("LAPSBarcode")
$WPFLAPSBarcode.Content = "$LAPSEncode"
$WPFLAPSBarcode.FontFamily = "$PSScriptRoot/LibreBarcode128-Regular.ttf#Libre Barcode 128"

$window.ShowDialog()

exit