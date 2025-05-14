#!/usr/bin/env pwsh
#
# SerialPortMaster.ps1 - Universal Serial Port Tool
#
# This script handles serial port communications with configurable parameters
# Supports predefined configurations, command files, and interactive mode

param (
    [Parameter(Mandatory=$false)]
    [string]$PortName = "COM1",
    
    [Parameter(Mandatory=$false)]
    [int]$BaudRate = 9600,
    
    [Parameter(Mandatory=$false)]
    [string]$Parity = "None",
    
    [Parameter(Mandatory=$false)]
    [int]$DataBits = 8,
    
    [Parameter(Mandatory=$false)]
    [string]$StopBits = "One",
    
    [Parameter(Mandatory=$false)]
    [string]$WindowTitle = "Serial Port Master",
    
    [Parameter(Mandatory=$false)]
    [string]$CommandFile,
    
    [Parameter(Mandatory=$false)]
    [int]$CommandDelay = 1000,
    
    [Parameter(Mandatory=$false)]
    [switch]$Interactive,
    
    [Parameter(Mandatory=$false)]
    [string]$Preset,
    
    [Parameter(Mandatory=$false)]
    [switch]$RecursiveCommands,
    
    [Parameter(Mandatory=$false)]
    [switch]$h
)

# Help output
if ($h) {
    Write-Host "SerialPortMaster.ps1 - Universal Serial Port Tool" -ForegroundColor Cyan
    Write-Host "Usage: .\SerialPortMaster.ps1 [parameters]`n" -ForegroundColor Cyan
    
    Write-Host "Parameters:" -ForegroundColor Green
    Write-Host "  -PortName     : Serial port name (default: COM1)"
    Write-Host "  -BaudRate     : Baud rate (default: 9600)"
    Write-Host "  -Parity       : Parity (None, Even, Odd, Mark, Space) (default: None)"
    Write-Host "  -DataBits     : Data bits (5, 6, 7, 8) (default: 8)"
    Write-Host "  -StopBits     : Stop bits (One, Two, OnePointFive) (default: One)"
    Write-Host "  -WindowTitle  : Terminal window title (default: Serial Port Master)"
    Write-Host "  -CommandFile  : File containing commands to send"
    Write-Host "  -CommandDelay : Delay between commands in milliseconds (default: 1000)"
    Write-Host "  -Interactive  : Enable interactive mode"
    Write-Host "  -RecursiveCommands : Loop through commands file repeatedly"
    Write-Host "  -Preset       : Use predefined configuration (Options: Sniffer, EnergyMeter, Default, RFEgypt)"
    Write-Host "  -h            : Display this help message`n"
    
    Write-Host "Presets:" -ForegroundColor Green
    Write-Host "  Sniffer  : 115200 baud, 8 data bits, No parity, 1 stop bit"
    Write-Host "  EnergyMeter  : 9600 baud, 7 data bits, Even parity, 1 stop bit" 
    Write-Host "  RFEgypt : 38400 baud, 7 data bits, Even parity, 1 stop bit`n"
    Write-Host "  Default : 9600 baud, 8 data bits, No parity, 1 stop bit"

    
    Write-Host "Special Characters:" -ForegroundColor Green
    Write-Host "  In command files, you can use the following escape sequences:"
    Write-Host "  \x02 - STX (Start of Text)"
    Write-Host "  \x03 - ETX (End of Text)"
    Write-Host "  \r   - CR (Carriage Return)"
    Write-Host "  \n   - LF (Line Feed)"
    Write-Host "  \x1B - ESC (Escape)"
    
    exit
}

# Apply preset configurations if specified
switch ($Preset) {
    "Sniffer" {
        $BaudRate = 115200
        $DataBits = 8
        $Parity = "None"
        $StopBits = "One"
        Write-Host "Using Sniffer preset configuration." -ForegroundColor Yellow
    }
    "EnergyMeter" {
        $BaudRate = 9600
        $DataBits = 7
        $Parity = "Even"
        $StopBits = "One"
        Write-Host "Using EnergyMeter preset configuration." -ForegroundColor Yellow
    }
    "Default" {
        $BaudRate = 9600
        $DataBits = 8
        $Parity = "None"
        $StopBits = "One"
        Write-Host "Using Default preset configuration." -ForegroundColor Yellow
    }
    "RFEgypt" {
        $BaudRate = 38400
        $DataBits = 7
        $Parity = "Even"
        $StopBits = "One"
        Write-Host "Using RFEgypt preset configuration." -ForegroundColor Yellow
    }
}

# Set the window title
$host.UI.RawUI.WindowTitle = $WindowTitle

# Function to parse special characters in command strings
function Parse-SpecialCharacters {
    param (
        [string]$InputString
    )
    
    $result = $InputString -replace '\\r', "`r" `
                          -replace '\\n', "`n" `
                          -replace '\\x02', [char]0x02 `
                          -replace '\\x03', [char]0x03 `
                          -replace '\\x1B', [char]0x1B
    
    return $result
}

# Function to display data with colorized special characters
function Display-ColorizedData {
    param (
        [string]$Data
    )
    
    # Only process if there's actual data to display
    if ([string]::IsNullOrEmpty($Data)) {
        return
    }
    
    for ($i = 0; $i -lt $Data.Length; $i++) {
        $char = $Data[$i]
        $byte = [byte][char]$char
        
        # Display control characters with special coloring
        if ($byte -lt 32 -or $byte -eq 127) {
            $ctrlChar = switch ($byte) {
                0 { "[NUL]" }
                1 { "[SOH]" }
                2 { "[STX]" }
                3 { "[ETX]" }
                4 { "[EOT]" }
                5 { "[ENQ]" }
                6 { "[ACK]" }
                7 { "[BEL]" }
                8 { "[BS]" }
                9 { "[HT]" }
                10 { "[LF]" }
                11 { "[VT]" }
                12 { "[FF]" }
                13 { "[CR]" }
                14 { "[SO]" }
                15 { "[SI]" }
                16 { "[DLE]" }
                17 { "[DC1]" }
                18 { "[DC2]" }
                19 { "[DC3]" }
                20 { "[DC4]" }
                21 { "[NAK]" }
                22 { "[SYN]" }
                23 { "[ETB]" }
                24 { "[CAN]" }
                25 { "[EM]" }
                26 { "[SUB]" }
                27 { "[ESC]" }
                28 { "[FS]" }
                29 { "[GS]" }
                30 { "[RS]" }
                31 { "[US]" }
                127 { "[DEL]" }
                default { $char }
            }
            # Use Magenta for special characters to make them stand out
            Write-Host $ctrlChar -ForegroundColor Magenta -NoNewline
            if ($ctrlChar -eq "[LF]") {
                Write-Host ""
            }
        } else {
            # Regular text in Cyan
            Write-Host $char -ForegroundColor Cyan -NoNewline
        }
    }
    # End with a newline
    Write-Host ""
    
    # Force garbage collection to reduce memory usage
    [System.GC]::Collect()
}

# Function to format and display special characters with color
function Format-DisplayData {
    param (
        [string]$Data
    )
    
    $result = ""
    
    for ($i = 0; $i -lt $Data.Length; $i++) {
        $char = $Data[$i]
        $byte = [byte][char]$char
        
        # Format control characters for display
        if ($byte -lt 32 -or $byte -eq 127) {
            $ctrlChar = switch ($byte) {
                0 { "[NUL]" }
                1 { "[SOH]" }
                2 { "[STX]" }
                3 { "[ETX]" }
                4 { "[EOT]" }
                5 { "[ENQ]" }
                6 { "[ACK]" }
                7 { "[BEL]" }
                8 { "[BS]" }
                9 { "[HT]" }
                10 { "[LF]" }
                11 { "[VT]" }
                12 { "[FF]" }
                13 { "[CR]" }
                14 { "[SO]" }
                15 { "[SI]" }
                16 { "[DLE]" }
                17 { "[DC1]" }
                18 { "[DC2]" }
                19 { "[DC3]" }
                20 { "[DC4]" }
                21 { "[NAK]" }
                22 { "[SYN]" }
                23 { "[ETB]" }
                24 { "[CAN]" }
                25 { "[EM]" }
                26 { "[SUB]" }
                27 { "[ESC]" }
                28 { "[FS]" }
                29 { "[GS]" }
                30 { "[RS]" }
                31 { "[US]" }
                32 { " " }
                127 { "[DEL]" }
                default { $char }
            }
            $result += $ctrlChar
        } else {
            $result += $char
        }
    }
    return $result
}

# Function to execute commands from file with memory optimization
function Execute-CommandFile {
    param (
        [string]$FilePath,
        [System.IO.Ports.SerialPort]$SerialPort,
        [int]$Delay
    )
    
    # Read all commands from file
    $commands = Get-Content $FilePath
    
    # Process each command
    foreach ($cmd in $commands) {
        # Skip empty lines and comments
        if ([string]::IsNullOrWhiteSpace($cmd) -or $cmd.Trim().StartsWith("#")) {
            continue
        }
        
        $parsedCmd = Parse-SpecialCharacters -InputString $cmd
        $SerialPort.Write($parsedCmd)
        Write-Host "Sent: $cmd" -ForegroundColor Yellow
        Start-Sleep -Milliseconds $Delay
        
        # Read any response
        try {
            $data = $SerialPort.ReadExisting()
            if (-not [string]::IsNullOrEmpty($data)) {
                Display-ColorizedData -Data $data
            }
        } catch [TimeoutException] {
            # Ignore timeout exceptions
        }
        
        # Clear variables to help with memory
        $parsedCmd = $null
        $data = $null
    }
    
    # Force garbage collection
    [System.GC]::Collect()
}

# Configure serial port
$port = $null
try {
    # Convert string parameters to appropriate .NET types
    $parityValue = [System.IO.Ports.Parity]::$Parity
    $stopBitsValue = [System.IO.Ports.StopBits]::$StopBits
    
    $port = New-Object System.IO.Ports.SerialPort $PortName, $BaudRate, $parityValue, $DataBits, $stopBitsValue
    $port.ReadTimeout = 1000  # Set a timeout for read operations in milliseconds
    
    # Open the serial port
    $port.Open()
    Write-Host "Serial port $($port.PortName) opened with settings: $BaudRate,$DataBits,$Parity,$StopBits" -ForegroundColor Green

    # Process command file if specified
    if ($CommandFile -and (Test-Path $CommandFile)) {
        Write-Host "Processing commands from file: $CommandFile" -ForegroundColor Yellow
        
        if ($RecursiveCommands) {
            Write-Host "Recursive command mode enabled. Will continuously loop through commands. Press Ctrl+C to stop." -ForegroundColor Green
            
            # Set counter to trigger garbage collection
            $loopCounter = 0
            
            # Loop indefinitely through the command file
            while ($true) {
                Execute-CommandFile -FilePath $CommandFile -SerialPort $port -Delay $CommandDelay
                Write-Host "Reached end of command file, restarting from beginning..." -ForegroundColor Yellow
                
                # Force garbage collection periodically
                $loopCounter++
                if ($loopCounter % 5 -eq 0) {
                    [System.GC]::Collect()
                    $loopCounter = 0
                }
            }
        } else {
            # Execute commands once
            Execute-CommandFile -FilePath $CommandFile -SerialPort $port -Delay $CommandDelay
        }
    }
    
    # Interactive mode
    if ($Interactive) {
        Write-Host "Interactive mode enabled. Type your command and press Enter to send." -ForegroundColor Green
        Write-Host "Press ESC to exit interactive mode." -ForegroundColor Yellow
        
        $inputBuffer = ""
        $promptShown = $false
        $cycleCounter = 0
        
        while ($true) {
            # Check for port data first
            if ($port.BytesToRead -gt 0) {
                $data = $port.ReadExisting()
                if (-not [string]::IsNullOrEmpty($data)) {
                    Display-ColorizedData -Data $data
                    $promptShown = $false
                }
                $data = $null
            }
            
            # Show prompt if needed
            if (-not $promptShown) {
                Write-Host "> " -NoNewline -ForegroundColor Green
                $promptShown = $true
            }
            
            # Check for input
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                
                # Check for escape key to exit
                if ($key.Key -eq [ConsoleKey]::Escape) {
                    Write-Host "`nExiting interactive mode." -ForegroundColor Yellow
                    $port.Close()
                    break
                }
                
                # Handle Enter key for command execution
                if ($key.Key -eq [ConsoleKey]::Enter) {
                    if ($inputBuffer.Length -gt 0) {
                        Write-Host ""  # New line after input
                        $parsedInput = Parse-SpecialCharacters -InputString $inputBuffer
                        $port.Write($parsedInput)
                        Write-Host "Sent: $inputBuffer" -ForegroundColor Yellow
                        $inputBuffer = ""
                        $parsedInput = $null
                        $promptShown = $false
                    }
                }
                # Handle Backspace
                elseif ($key.Key -eq [ConsoleKey]::Backspace) {
                    if ($inputBuffer.Length -gt 0) {
                        $inputBuffer = $inputBuffer.Substring(0, $inputBuffer.Length - 1)
                        Write-Host "`b `b" -NoNewline  # Erase character on screen
                    }
                }
                # Regular character input
                else {
                    $inputBuffer += $key.KeyChar
                    Write-Host $key.KeyChar -NoNewline
                }
            }
            
            # Small sleep to prevent CPU hogging
            Start-Sleep -Milliseconds 100
            
            # Periodically force garbage collection
            $cycleCounter++
            if ($cycleCounter -ge 100) {
                [System.GC]::Collect()
                $cycleCounter = 0
            }
        }
    }
    
    # Default listening mode if no interactive or command file was specified,
    # or if command file was processed without recursion
    if ((-not $Interactive -and -not $CommandFile) -or 
        ($CommandFile -and -not $RecursiveCommands) -or
        ($Interactive -and -not $port.IsOpen)) {  # Fall back if interactive mode was exited
        Write-Host "Listening on port $($port.PortName)... Press Ctrl+C to stop." -ForegroundColor Green
        if (($Interactive -and -not $port.IsOpen)) {
            $port.Open()
        }
        
        # Counter for garbage collection
        $listenCounter = 0
        
        # Continuously read data
        while ($true) {
            try {
                $data = $port.ReadExisting()
                if (-not [string]::IsNullOrEmpty($data)) {
                    Display-ColorizedData -Data $data
                    $data = $null
                }
                Start-Sleep -Milliseconds 1000  # Longer sleep for better memory usage
                
                # Periodically force garbage collection
                $listenCounter++
                if ($listenCounter -ge 10) {
                    [System.GC]::Collect()
                    $listenCounter = 0
                }
            } catch [TimeoutException] {
                # Ignore timeout exceptions
            }
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Close the port and clean up
    if ($port -and $port.IsOpen) {
        $port.Close()
        Write-Host "Port closed." -ForegroundColor Yellow
    }
    # Final garbage collection before exit
    [System.GC]::Collect()
} 