# Serial Port Master

A universal PowerShell script for serial port communication that combines listening, command sending, and interactive modes.

## Features

- Configurable serial port settings (port, baud rate, parity, data bits, stop bits)
- Flow control support: RTS/CTS (hardware) and XON/XOFF (software), independently or combined
- Configurable write timeout to prevent hangs when flow control stalls
- Predefined configurations for common scenarios
- Send commands from a file with configurable delay
- Recursive command execution mode
- Interactive mode for manual command entry
- Special character support (STX, ETX, CR, LF, etc.)
- Custom window title for multiple instances
- Comprehensive logging of all communication with timestamps
- Automatic log file rotation to prevent excessive disk usage

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
| -RtsCts            | Enable RTS/CTS hardware flow control            | False          |
| -XonXoff           | Enable XON/XOFF software flow control           | False          |
| -WriteTimeout      | Write timeout in ms (use -1 for infinite)       | 5000           |
| -WindowTitle       | Terminal window title                           | Serial Port Master |
| -CommandFile       | File containing commands to send                |                |
| -CommandDelay      | Delay between commands in milliseconds          | 1000           |
| -Interactive       | Enable interactive mode                         | False          |
| -RecursiveCommands | Loop through commands file repeatedly           | False          |
| -LogFile           | File to log all sent and received data          |                |
| -MaxLogSizeMB      | Maximum log file size in MB before rotation     | 10             |
| -Preset            | Use predefined configuration                    |                |
| -h                 | Display help message                            |                |

### Preset Configurations

- **Sniffer**: 115200 baud, 8 data bits, No parity, 1 stop bit
- **EnergyMeter**: 9600 baud, 7 data bits, Even parity, 1 stop bit
- **RFEgypt**: 38400 baud, 7 data bits, Even parity, 1 stop bit
- **Default**: 9600 baud, 8 data bits, No parity, 1 stop bit

Presets do not set flow control. Combine a preset with `-RtsCts` and/or `-XonXoff` if the target device requires it.

### Flow Control

`-RtsCts` and `-XonXoff` are independent switches and map to `System.IO.Ports.Handshake` as follows:

| `-RtsCts` | `-XonXoff` | Handshake mode         |
|-----------|------------|------------------------|
| off       | off        | None (default)         |
| on        | off        | RequestToSend          |
| off       | on        | XOnXOff                |
| on        | on        | RequestToSendXOnXOff   |

The resolved handshake mode is printed on startup and written to the log header.

#### Things to be aware of

- **CTS must be asserted for writes to proceed with `-RtsCts`.** If the peer never raises CTS, every `Write()` would block forever. `-WriteTimeout` (default 5000 ms) bounds this: a timed-out write is logged as an `ERROR` entry and the session continues instead of hanging. Set `-WriteTimeout -1` to restore the previous infinite-wait behaviour, or increase it for slow peers.
- **XON/XOFF consumes 0x11 (DC1) and 0x13 (DC3).** When `-XonXoff` is on, those two bytes are intercepted by the driver as flow-control signals and never reach the receive side, so `[DC1]` / `[DC3]` will not appear in the console output or log for that session. This is driver behaviour, not a bug.
- **Flow control cannot be changed after the port is open** — stop the script and relaunch with different switches to change modes.

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
.\SerialPortMaster.ps1 -PortName COM5 -Preset Sniffer -Interactive
```

### Enabling Flow Control

```powershell
# Hardware flow control (RTS/CTS) with a 10-second write timeout
.\SerialPortMaster.ps1 -PortName COM3 -BaudRate 115200 -RtsCts -WriteTimeout 10000

# Software flow control (XON/XOFF)
.\SerialPortMaster.ps1 -PortName COM3 -XonXoff

# Both hardware and software flow control together
.\SerialPortMaster.ps1 -PortName COM3 -RtsCts -XonXoff
```

### Interactive Mode with Logging

```powershell
.\SerialPortMaster.ps1 -PortName COM5 -Preset Sniffer -Interactive -LogFile "serial_log.txt"
```

### Logging with Custom Maximum Log Size

```powershell
.\SerialPortMaster.ps1 -PortName COM5 -Preset Sniffer -Interactive -LogFile "serial_log.txt" -MaxLogSizeMB 50
```

### Using Multiple Instances

You can run multiple instances with different window titles:

```powershell
# First terminal
.\SerialPortMaster.ps1 -PortName COM3 -WindowTitle "RF Receiver" -LogFile "receiver_log.txt"

# Second terminal
.\SerialPortMaster.ps1 -PortName COM4 -WindowTitle "RF Transmitter" -CommandFile rf_commands.txt -RecursiveCommands -LogFile "transmitter_log.txt"
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

## Log File Format

When the `-LogFile` parameter is used, the script creates a detailed log with the following information:

- Timestamps with millisecond precision
- Direction indicators (SENT, RECV, INFO, ERROR)
- Raw data with special characters represented in a readable format
- Session start/end markers with configuration details

This is especially useful for debugging communication issues or creating audit trails of device interactions.

### Log Rotation

The script automatically manages log file size by rotating logs when they exceed the specified size (default: 10MB). When a log file reaches the maximum size:

1. The current log file is renamed with a timestamp (e.g., `serial_log.txt.20230830_123456.bak`)
2. A new log file is created to continue logging
3. A rotation message is written to both console and the new log file

This prevents logs from consuming excessive disk space during long-running sessions or high-volume communications. 