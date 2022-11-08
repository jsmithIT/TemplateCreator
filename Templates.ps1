##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTkaDjpbH6iJy91iuQ2YvLvaUv6KwxZO5w9nvqSzJXYkoYGA3uS/6DUi4GcYXWOUZtcUURiEYKuEE96DTHvSVUacHgu9+f+DOr7EmGxTR4JyU
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTkaDjpbH6iJy91iuQ2YvLvaUv6KwxZO5w9nvqSzJXYkoYGA3uS/6DUi4GcYXWOUZtcUURiEDG7Ik5LTZGOLpcacHgu9+f+Cd6LcxEDo=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
.SYNOPSIS
A powershell script to standardize documentation by creating County departmental templates.

.DESCRIPTION
This powershell script will quietly add the templates to an Office template directory, and direct the relevant Office applications to said directory.
Authors: JBSMITH
Version: 1.0.5

#>

##* ANCHOR: Global Variables

# Registry key path for Word template
$KeyPathWord = "HKCU:\Software\Microsoft\Office\16.0\word\options"

# Registry key path for Excel template
$KeyPathExcel = "HKCU:\Software\Microsoft\Office\16.0\excel\options"

# Registry key path for Powerpoint template
$KeyPathPP = "HKCU:\Software\Microsoft\Office\16.0\powerpoint\options"

# Key value/path to Office templates
$ValuePath = "$env:USERPROFILE\Documents\Custom Office Templates"

# Folder path for templates to be added.
$FolderPath = "S:\Drivers and Software\Scripts\IT Office Templates\Files\Custom Office Templates"

# Key value name
$Value = "PersonalTemplates"

##* ANCHOR: Add Templates Function
function regCheck() {
    # If Word has never been initialized, create new registry key.
    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\word")) {
        Try {
            Write-Host -Message "Creating Word registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0" -Name "word"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Word registry key."
        }
    }
    Else {
        Write-Host "Word registry key exists."
    }

    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\word\options")) {
        Try {
            Write-Host -Message "Creating Word registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\word" -Name "options"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Word registry key."
        }
    }
    Else {
        Write-Host "Word has been previously initialized."
    }

    # If Excel has never been initialized, create new registry key.
    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\excel")) {
        Try {
            Write-Host -Message "Creating Excel registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0" -Name "excel"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Excel registry key."
        }
    }
    Else {
        Write-Host "Excel registry key exists."
    }

    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\excel\options")) {
        Try {
            Write-Host -Message "Creating Excel registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\excel" -Name "options"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Excel registry key."
        }
    }
    Else {
        Write-Host "Excel has been previously initialized."
    }

    # If Powerpoint has never been initialized, create new registry key.
    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\powerpoint")) {
        Try {
            Write-Host -Message "Creating Powerpoint registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0" -Name "powerpoint"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Powerpoint registry key."
        }
    }
    Else {
        Write-Host "Powerpoint registry key exists."
    }

    If (-not (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\powerpoint\options")) {
        Try {
            Write-Host -Message "Creating Powerpoint registry key."
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\powerpoint" -Name "options"
        }
        Catch [System.Exception] {
            Write-Host "Failed to create Powerpoint registry key."
        }
    }
    Else {
        Write-Host "Powerpoint has been previously initialized."
    }
}

##* ANCHOR: Add Templates Function
function templateImplementation() {

    # Adds templates and specifies directory.
    Try {
        # Changes registry keys for Office applications to point to the template directory.
        Write-Host "Adjusting registry keys for template directories."
        New-ItemProperty -Path "$KeyPathWord" -Name "$Value" -Value "$ValuePath" -PropertyType STRING -Force 
        New-ItemProperty -Path "$KeyPathExcel" -Name "$Value" -Value "$ValuePath" -PropertyType STRING -Force
        New-ItemProperty -Path "$KeyPathPP" -Name "$Value" -Value "$ValuePath" -PropertyType STRING -Force

        # Checked registry value for Word
        $MainKeyPath1 = Get-ItemPropertyValue -Path "$KeyPathWord" -Name "$Value"
        # Checked registry value for Excel
        $MainKeyPath2 = Get-ItemPropertyValue -Path "$KeyPathExcel" -Name "$Value"
        # Checked registry value for Powerpoint
        $MainKeyPath3 = Get-ItemPropertyValue -Path "$KeyPathPP" -Name "$Value"
        
        # If the key exists, log that it was successful. 
        If ($MainKeyPath1 -eq "$ValuePath" -and $MainKeyPath2 -eq "$ValuePath" -and $MainKeyPath3 -eq "$ValuePath") {
            Write-Host "Key changes completed successfully."
        }
        # Otherwise, log that it was unsuccessful.
        Else {
            Write-Host "Key changes failed to complete."
        }
    }
    # Log if the registry key does not exist.
    Catch [PathNotFound, Microsoft.PowerShell.Commands.NewItemPropertyCommand] {
        Write-Host "The specified registry key does not exist."
    }
    # Log if the changes fail.
    Catch [System.Exception] {
        Write-Host "Key changes failed to complete."
    }

    If (-not(Test-Path "$ValuePath")) {
        # Adds templates to directory.
        Try {
            # Creates a folder for Office templates and puts all standardized templates into it.
            Write-Host "Adding templates to directory."
            Copy-Item -Path "$FolderPath" -Destination "$ValuePath" -Recurse -Force

            # If the key exists, log that it was successful. 
            If (Test-Path -Path "$ValuePath") {
                Write-Host "Templates added successfully."
            }
            # Otherwise, log that it was unsuccessful.
            Else {
                Write-Host "Templates failed to add."
            }
        }
        # Log if the files do not exist.
        Catch [Microsoft.PowerShell.Commands.CopyItemCommand] {
            Write-Host "Cannot find files in script root."
        }
        # Log if templates do not add.
        Catch [System.Exception] {
            Write-Host "Templates failed to add."
        }
    }
    Else {
        # Adds templates to directory.
        Try {
            # Puts all standardized templates into the template directory.
            Write-Host "Adding templates to directory."
            Copy-Item -Path "$FolderPath\*" -Destination "$ValuePath" -Recurse -Force -PassThru

            # If the key exists, log that it was successful. 
            If (Test-Path -Path "$ValuePath") {
                Write-Host "Templates added successfully."
            }
            # Otherwise, log that it was unsuccessful.
            Else {
                Write-Host "Templates failed to add."
            }
        }
        # Log if the files do not exist.
        Catch [Microsoft.PowerShell.Commands.CopyItemCommand] {
            Write-Host "Cannot find files in script root."
        }
        # Log if templates do not add.
        Catch [System.Exception] {
            Write-Host "Templates failed to add."
        }
    }
}

# Calls the function to initialize the applications if necessary.
regCheck

#Pauses the script for one second to alow it to recognize the key changes.
Start-Sleep -s 1

# Calls the function to add the templates, changing the registries in the process if all registry keys exist.
If ((Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\word") -and (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\excel") -and (Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\powerpoint")) {
    templateImplementation
    
}