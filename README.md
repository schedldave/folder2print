# Folder2Print

Automatically print PDF files from a watched folder. When a PDF file is synchronized/copied to the configured folder, it will be sent to the printer automatically.

## Features

- **Folder Monitoring**: Watches a folder for new PDF files
- **Automatic Printing**: Sends new files to a configured printer
- **Sync-Safe**: Waits for files to be fully written before printing
- **File Management**: Option to move or delete files after printing
- **Configurable**: JSON configuration file for all settings
- **Logging**: Detailed logging to console and file

## Requirements

- Windows 10/11
- Python 3.8+
- Adobe Acrobat or Acrobat Reader (recommended for `acrobat` print method)

## Installation

1. **Install Python dependencies:**

   ```cmd
   pip install -r requirements.txt
   ```

2. **List available printers:**

   ```cmd
   python folder2print.py --list-printers
   ```

3. **Edit `config.json`** with your settings:
   - `watch_folder`: The folder to monitor for PDF files
   - `printer_name`: The printer to use (leave empty for default printer)

## Configuration

Edit `config.json`:

```json
{
    "watch_folder": "C:\\Sync\\PDFs",
    "printer_name": "HP LaserJet Pro",
    "check_interval_seconds": 1,
    "delete_after_print": false,
    "move_after_print": true,
    "printed_folder": "printed",
    "file_extensions": [".pdf"],
    "print_delay_seconds": 2,
    "move_delete_delay_seconds": 60,
    "print_method": "acrobat",
    "acrobat_path": "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe"
}
```

### Configuration Options

| Option | Description | Default |
|--------|-------------|---------|
| `watch_folder` | Folder path to monitor for new files | Required |
| `printer_name` | Printer name (use `--list-printers` to see options) | Default printer |
| `check_interval_seconds` | How often to check for changes | 1 |
| `delete_after_print` | Delete files after printing | false |
| `move_after_print` | Move files to subfolder after printing | true |
| `printed_folder` | Subfolder name for printed files | "printed" |
| `file_extensions` | File types to monitor | [".pdf"] |
| `print_delay_seconds` | Extra delay before printing (for sync) | 2 |
| `move_delete_delay_seconds` | Delay before moving/deleting after print | 60 |
| `print_method` | Printing method: `acrobat` or `shellexecute` | "acrobat" |
| `acrobat_path` | Path to Acrobat executable (auto-detected if empty) | Auto-detect |

### Print Methods

**`acrobat` (recommended)**: Uses Adobe Acrobat/Reader command line (`/t` switch) for reliable silent printing. The path is auto-detected from common installation locations, or you can specify it manually.

Common Acrobat paths:
- Acrobat DC: `C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe`
- Acrobat Reader DC: `C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe`

**`shellexecute`**: Uses Windows ShellExecute with the "print" verb. Works with any PDF application but may show dialogs.

## Usage

### Run manually:

```cmd
python folder2print.py
```

### Run with custom config:

```cmd
python folder2print.py myconfig.json
```

### List available printers:

```cmd
python folder2print.py --list-printers
```

### Run at Windows Startup (Autostart)

**Option 1: Startup Folder**

1. Press `Win + R`, type `shell:startup`, press Enter
2. Create a shortcut to `run_folder2print.bat` in that folder

**Option 2: Task Scheduler**

1. Open Task Scheduler
2. Create Basic Task
3. Trigger: "When I log on"
4. Action: Start a program
5. Program: `pythonw.exe`
6. Arguments: `"C:\GitHub\folder2print\folder2print.py"`
7. Start in: `C:\GitHub\folder2print`

**Option 3: Using pythonw.exe (no console window)**

```cmd
pythonw.exe folder2print.py
```

## Logging

Logs are written to:
- Console output (when running manually)
- `folder2print.log` file in the script directory

## Troubleshooting

### "No printers found"
- Ensure you have at least one printer installed
- Run `--list-printers` to verify

### "File not ready after timeout"
- The file may still be syncing
- Increase `print_delay_seconds` in config

### "Acrobat executable not found"
- Install Adobe Acrobat or Acrobat Reader
- Set `acrobat_path` in config.json to the correct path
- Or switch to `"print_method": "shellexecute"`

### PDF doesn't print
- Check the log file `folder2print.log` for errors
- Verify Acrobat can open the PDF manually
- Try switching `print_method` between `acrobat` and `shellexecute`

### Files not detected
- Check that `watch_folder` path is correct
- Ensure the folder exists
- Check file extension matches `file_extensions` config

### Print job sent but nothing prints
- Increase `move_delete_delay_seconds` to give more time for printing
- Check printer queue in Windows settings
- Verify printer is online and has paper

## License

MIT License
