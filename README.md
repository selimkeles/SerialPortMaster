# Serial Port Master

A universal PowerShell script for serial port communication that combines listening, command sending, and interactive modes.

## Features

- Configurable serial port settings (port, baud rate, parity, data bits, stop bits)
- Predefined configurations for common scenarios
- Send commands from a file with configurable delay
- Recursive command execution mode
- Interactive mode for manual command entry
- Special character support (STX, ETX, CR, LF, etc.)
- Custom window title for multiple instances

## Usage

```powershell
.\SerialPortMaster.ps1 [parameters]
```

### Parameters

| Parameter          | Description                                     | Default        |
|--------------------|-------------------------------------------------|----------------|
| -PortName          | Serial port name                                | COM1           |
| -BaudRate          | Baud rate                                       | 9600           |
| -Parity            | Parity (None, Even, Odd, Mark, Space)           | None           |
| -DataBits          | Data bits (5, 6, 7, 8)                          | 8              |
| -StopBits          | Stop bits (One, Two, OnePointFive)              | One            |
| -WindowTitle       | Terminal window title                           | Serial Port Master |
| -CommandFile       | File containing commands to send                |                |
| -CommandDelay      | Delay between commands in milliseconds          | 1000           |
| -Interactive       | Enable interactive mode                         | False          |
| -RecursiveCommands | Loop through commands file repeatedly           | False          |
| -Preset            | Use predefined configuration                    |                |
| -h                 | Display help message                            |                |

### Preset Configurations

- **Sniffer**: 115200 baud, 8 data bits, No parity, 1 stop bit
- **EnergyMeter**: 9600 baud, 7 data bits, Even parity, 1 stop bit
- **RFEgypt**: 38400 baud, 7 data bits, Even parity, 1 stop bit
- **Default**: 9600 baud, 8 data bits, No parity, 1 stop bit

## Examples

### Simple Listening Mode

```powershell
.\SerialPortMaster.ps1 -PortName COM3 -WindowTitle "Device Monitor"
```

### Send Commands from File

```powershell
.\SerialPortMaster.ps1 -PortName COM4 -CommandFile commands.txt -CommandDelay 2000
```

### Send Commands Repeatedly from File

```powershell
.\SerialPortMaster.ps1 -PortName COM4 -CommandFile commands.txt -CommandDelay 2000 -RecursiveCommands
```

### Interactive Mode with Preset Configuration

```powershell
.\SerialPortMaster.ps1 -PortName COM5 -Preset Sniffer115200 -Interactive
```

### Using Multiple Instances

You can run multiple instances with different window titles:

```powershell
# First terminal
.\SerialPortMaster.ps1 -PortName COM3 -WindowTitle "RF Receiver"

# Second terminal
.\SerialPortMaster.ps1 -PortName COM4 -WindowTitle "RF Transmitter" -CommandFile rf_commands.txt -RecursiveCommands
```

## Command File Format

The command file contains one command per line. Lines starting with # are treated as comments.
Special characters can be included using escape sequences:

- `\x02` - STX (Start of Text)
- `\x03` - ETX (End of Text)
- `\r` - CR (Carriage Return)
- `\n` - LF (Line Feed)
- `\x1B` - ESC (Escape)

Example command file (`example_commands.txt`):
```
# RF commands
+++\r\n
AT+MODE=4\r\n

# Command with special characters
\x02DATA REQUEST\x03\r\n
``` 