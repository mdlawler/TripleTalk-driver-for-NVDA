# Add kernel32.dll functions for working with INI files
Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class IniFile {
[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
public static extern int GetPrivateProfileString(string section, string key, string defaultValue, StringBuilder returnValue, int size, string filePath);

[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
public static extern bool WritePrivateProfileString(string section, string key, string value, string filePath);
}
"@ -PassThru

# Ensure a command-line argument is provided
if ($args.Count -ne 1 -or ($args[0] -ne "0" -and $args[0] -gt "3")) {
Write-Host "Usage: script.ps1 <0 to 3>"
Exit 1
}

# Argument: 1 or 0
$value = $args[0]

# Path to the INI file (in Windows directory)
$iniPath = "C:\Windows\ttusbd.ini"

# Section and key in the INI file
$section = "ttalk_usb_comm"
$key = "nopauses"

# Write the argument value to the INI file
if ([IniFile]::WritePrivateProfileString($section, $key, $value, $iniPath)) {
Write-Host "Successfully wrote '$value' to [$section] $key in $iniPath."
} else {
Write-Host "Failed to write to the INI file."
}
