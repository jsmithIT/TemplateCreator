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
##L8/UAdDXTkaDjpbH6iJy91iuQ2YvLvaUv6KwxZO5w9nvqSzJXYkoYGA3uS/6DUi4GcYXWOUZtcUURiEDG7Ik5LTZGOLpcacHgu9+f+Cd5ocnHFTL95L431eX15ig
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
.SYNOPSIS
A powershell script to remove standardized documentation.

.DESCRIPTION
This powershell script will quietly remove the templates from an Office template directory.
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

##* ANCHOR: Remove Templates Directory
function regCheck() {
    # If Templates has been installed, remove the registry keys that were created.
    If (Test-Path -Path "HKLM:\SOFTWARE\Tuolumne County\Package Information\$apName*") { 
        # Removes Word registry key.
        Try {
            Write-Host -Message "Removing Word registry key."
            Remove-Item -Path "$KeyPathWord"
        }
        Catch [System.Exception] {
            Write-Host "Failed to remove Word registry key."
        }

        # Removes Excel registry key.
        Try {
            Write-Host -Message "Removing Excel registry key."
            Remove-Item -Path "$KeyPathExcel"
        }
        Catch [System.Exception] {
            Write-Host "Failed to remove Excel registry key."
        }

        # Removes Powerpoint registry key.
        Try {
            Write-Host -Message "Removing Powerpoint registry key."
            Remove-Item -Path "$KeyPathPP"
        }
        Catch [System.Exception] {
            Write-Host "Failed to remove Powerpoint registry key."
        }
    }
    Else {
        Write-Host "Registry keys do not exist."
    }
}

    ##* ANCHOR: Remove Templates Function
    function templateRemoval() {
        If ((Test-Path "$ValuePath")) {
            # Removes templates to directory.
            Try {
                # Removes the template directory folder.
                Write-Host "Removing templates from directory."
                Remove-Item -Path "$ValuePath" -Force -Recurse

                # If the location does not exist, log that it was successful. 
                If (-not(Test-Path -Path "$ValuePath")) {
                    Write-Host "Templates removed successfully."
                }
                # Otherwise, log that it was unsuccessful.
                Else {
                    Write-Host "Templates failed to remove."
                }
            }
            # Log if templates do not remove.
            Catch [System.Exception] {
                Write-Host "Templates failed to remove."
            }
        }
        Else {
            Write-Host "The templates do not exist."
        }
    }

    # Calls the function to remove the registry keys.
    regCheck

    #Pauses the script for one second to alow it to recognize the key changes.
    Start-Sleep -s 1

    # Calls the function to remove the templates, changing the registries in the process if all registry keys exist.
    templateRemoval