# Clear the console at the beginning
Clear-Host

# Banner
@'

   ___      _      __          __  ___                                       __    
  / _ \____(_)__  / /____ ____/  |/  /__ ____  ___ ____ ____ __ _  ___ ___  / /_   
 / ___/ __/ / _ \/ __/ -_) __/ /|_/ / _ `/ _ \/ _ `/ _ `/ -_)  ' \/ -_) _ \/ __/   
/_/  /_/ /_/_//_/\__/\__/_/ /_/  /_/\_,_/_//_/\_,_/\_, /\__/_/_/_/\__/_//_/\__/    
                                                  /___/                                                                                         
Version: 2.5                                            
'@

# Function to handle errors
function Handle-Error {
    param (
        [string]$errorMessage
    )
    Write-Host "Error: $errorMessage" -ForegroundColor Red
}

function Validate-FilePath {
    param (
        [string]$filePath
    )

    while (-not (Test-Path $filePath -PathType Leaf) -or [string]::IsNullOrWhiteSpace($filePath)) {
        Write-Warning "The file path is invalid or empty. Please try again."
        $filePath = Read-Host "Enter the file path"
    }

    return $filePath
}

function Add-Printers {
    # Script title
    Write-Host "Adding printers by local port and IP"

    # Show current folder location
    Get-ChildItem

    # Loop until a valid printer file path is entered
    $printersFile = Validate-FilePath -filePath (Read-Host "Enter the printer file path")

    # Import printers file in CSV format with ';' as delimiter
    $printerList = Import-Csv $printersFile -Delimiter ';'

    # Show start message
    Write-Host "Starting..."

    # Create printers specified in the printers file
    foreach ($printer in $printerList) {
        Write-Host "Creating printer $($printer.Name) on port $($printer.LocalPort)"

        # Check if printer port already exists, if not, create it
        $portExists = Get-PrinterPort -Name $printer.LocalPort -ErrorAction SilentlyContinue
        if (-not $portExists) {
            try {
                if ($printer.LocalPort -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") {
                    Add-PrinterPort -Name $printer.LocalPort -PrinterHostAddress $printer.LocalPort
                } else {
                    Add-PrinterPort -Name $printer.LocalPort
                }
            } catch {
                Handle-Error "Error creating port $($printer.LocalPort): $_"
                continue  # Jump to the next iteration of the loop in case of error
            }
        }

        # Check if printer driver is already installed, if not, show a warning
        $printDriverExists = Get-PrinterDriver -Name $printer.Driver -ErrorAction SilentlyContinue
        if ($printDriverExists) {
            try {
                Add-Printer -Name $printer.Name -PortName $printer.LocalPort -DriverName $printer.Driver -ErrorAction Stop
                Write-Host "Printer $($printer.Name) created on port $($printer.LocalPort)"
            } catch {
                Handle-Error "Error adding printer $($printer.Name): $_"
                Write-Host "Error details:"
                Write-Host $_.Exception.Message
                Write-Host $_.Exception.StackTrace
            }
        } else {
            Write-Warning "Printer driver $($printer.Name) is not installed."
        }
    }

    # Print a completion message
    Read-Host -Prompt "Printers created successfully. Press any key to extit..."
}

# Function to send test pages
function Send-TestPages {
    # Script title
    Write-Host "Sending test pages in bulk"

    # Show current folder location
    Get-ChildItem

    # Loop until a valid printer file path is entered
    # Ensure provided path exists and is not empty.
    do {
        $printersFile = Read-Host "Enter the printer file path"

        if (-not [string]::IsNullOrWhiteSpace($printersFile) -and -not (Test-Path $printersFile)) {
            Write-Warning "The file path is invalid. Please try again."
        }
    } until (-not [string]::IsNullOrWhiteSpace($printersFile) -and (Test-Path $printersFile))

    # Show a message indicating that the printers file is being read.
    Write-Host "Reading file $printersFile..."

    # Read the printers file and store printer names in the $printers variable.
    $printers = Import-Csv -Path $printersFile -Delimiter ';'

    # Send a test page to each printer.
    foreach ($printer in $printers) {
        # Show a message indicating that a test page is being sent to the printer.
        Write-Host "Sending test page to $($printer.Name) on port $($printer.LocalPort)..."
        
        # Execute the print command in the background
        $jobScript = {
            param ($printer, $env:COMPUTERNAME)
            $command = "rundll32.exe printui.dll,PrintUIEntry /k /n\\$env:COMPUTERNAME\$($printer.Name)"
            Start-Process cmd -ArgumentList "/c $command" -NoNewWindow -Wait
        }

        Start-Job -ScriptBlock $jobScript -ArgumentList $printer, $env:COMPUTERNAME | Wait-Job | Receive-Job

        Write-Host "Test page sent successfully to $($printer.Name) on port $($printer.LocalPort)."
    }

    # Show a confirmation message that test pages have been sent to all printers.
    Read-Host -Prompt "Completed sending test pages to all printers. Press any key to exit..."
}

# Function to remove printers
function Remove-Printers {
    # Script title
    Write-Host "Removing printers."

    # Show current folder location
    Get-ChildItem

    # Loop until a valid printer file path is entered
    $printersFile = Validate-FilePath -filePath (Read-Host "Enter the printer file path")

    # Import printers list from CSV file
    $printerList = Import-Csv $printersFile -Delimiter ';'

    # Start the process of removing printers and ports
    Write-Host "Starting..."

    foreach ($printer in $printerList) {
        # Remove the printer
        Write-Host "Removing printer $($printer.Name)"
        $printerExists = Get-Printer -Name $printer.Name -ErrorAction SilentlyContinue
        if ($printerExists) {
            Remove-Printer -Name $printer.Name
            Write-Host "Printer $($printer.Name) removed."
        } else {
            # Show a warning message if the printer does not exist
            Write-Warning "Printer $($printer.Name) does not exist."
        }

        # Remove the printer port
        Write-Host "Removing port $($printer.LocalPort)"
        $portExists = Get-PrinterPort -Name $printer.LocalPort -ErrorAction SilentlyContinue
        if ($portExists) {
            Remove-PrinterPort -Name $printer.LocalPort
            Write-Host "Port $($printer.LocalPort) removed."
        } else {
            # Show a warning message if the port does not exist
            Write-Warning "Port $($printer.LocalPort) does not exist."
        }
    }
    # Show a confirmation message that all the printers and ports have been removed.
    Read-Host -Prompt "All printers have been removed. Press any key to exit..."
}

# Function to clear print queues
function Clear-PrintQueues {
    # Script title
    Write-Host "Mass clearing of print queues"
    
    Stop-Service spooler
    Remove-Item -Path $env:windir\system32\spool\PRINTERS\*.*

    $start = Start-Service Spooler -ErrorAction Ignore

    if ((Get-Service spooler).status -eq 'Stopped') {
        Start-Service Spooler -ErrorAction Ignore
    }
    # Shows that the printer queue have been cleared.
    Read-Host -Prompt "Print queue cleared successfully! Press any key to exit..."
}

# Function to inventory printers
function Inventory-Printers {
    # Script title
    Write-Host "Inventorying printers..."

    # Logic to inventory printers
    Get-WmiObject -class win32_printer -ComputerName $env:COMPUTERNAME | Select Caption,PortName,DriverName,PrinterStatus | Export-Csv -Path .\inventory.csv -Delimiter ';' -NoTypeInformation
    
    # Shows that all the printers have been inventoried.
    Read-Host -Prompt "Printers inventoried successfully! Press any key to exit..."
}

# Main menu
Write-Host "1. Add Printers"
Write-Host "2. Remove Printers"
Write-Host "3. Send Test Pages"
Write-Host "4. Clear Print Queue"
Write-Host "5. Inventory Printers"
Write-Host "6. Exit"
# Validate user's option
do {
    $option = Read-Host "Choose an option (1-6)"
    if ($option -notmatch '^[1-6]$') {
        Write-Warning "Please enter a valid option (1-6)."
    }
} until ($option -match '^[1-6]$')

switch ($option) {
    "1" {
        # Call Add-Printers function
        Add-Printers
    }
    "2" {
        # Call Remove-Printers function
        Remove-Printers
    }
    "3" {
        # Call Send-TestPages function
        Send-TestPages
    }
    "4" {
        # Call Clear-PrintQueues function
        Clear-PrintQueues
    }
    "5" {
        # Call Inventory-Printers function
        Inventory-Printers
    }
    "6" {
        # Exit the program
        break
    }

    default {
        Write-Host "Invalid option"
    }
}
