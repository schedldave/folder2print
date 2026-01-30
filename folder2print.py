"""
Folder2Print - Automatically print PDF files from a watched folder.

This script monitors a specified folder for new PDF files and sends them
to a configured printer automatically. Designed for Windows.
"""

import json
import logging
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import win32print
import win32api
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileCreatedEvent


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("folder2print.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


class Config:
    """Configuration handler for the application."""

    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.watch_folder: str = ""
        self.printer_name: str = ""
        self.check_interval_seconds: int = 1
        self.delete_after_print: bool = False
        self.move_after_print: bool = True
        self.printed_folder: str = "printed"
        self.file_extensions: list = [".pdf"]
        self.print_delay_seconds: int = 2
        self.move_delete_delay_seconds: int = 60
        self.print_method: str = "acrobat"  # "acrobat" or "shellexecute"
        self.acrobat_path: str = ""

        self.load()

    def load(self) -> None:
        """Load configuration from JSON file."""
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.watch_folder = data.get("watch_folder", "")
            self.printer_name = data.get("printer_name", "")
            self.check_interval_seconds = data.get("check_interval_seconds", 1)
            self.delete_after_print = data.get("delete_after_print", False)
            self.move_after_print = data.get("move_after_print", True)
            self.printed_folder = data.get("printed_folder", "printed")
            self.file_extensions = data.get("file_extensions", [".pdf"])
            self.print_delay_seconds = data.get("print_delay_seconds", 2)
            self.move_delete_delay_seconds = data.get("move_delete_delay_seconds", 60)
            self.print_method = data.get("print_method", "acrobat")
            self.acrobat_path = data.get("acrobat_path", "")

            # Auto-detect Acrobat path if not specified
            if not self.acrobat_path:
                self.acrobat_path = self._find_acrobat_path()

            logger.info(f"Configuration loaded from {self.config_path}")

        except FileNotFoundError:
            logger.error(f"Configuration file not found: {self.config_path}")
            self.create_default_config()
            raise SystemExit(
                f"Please edit {self.config_path} and restart the application."
            )
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in configuration file: {e}")
            raise SystemExit("Configuration file contains invalid JSON.")

    def create_default_config(self) -> None:
        """Create a default configuration file."""
        default_config = {
            "watch_folder": "C:\\Sync\\PDFs",
            "printer_name": "",
            "check_interval_seconds": 1,
            "delete_after_print": False,
            "move_after_print": True,
            "printed_folder": "printed",
            "file_extensions": [".pdf"],
            "print_delay_seconds": 2,
            "move_delete_delay_seconds": 60,
            "print_method": "acrobat",
            "acrobat_path": "",
        }

        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4)

        logger.info(f"Default configuration created: {self.config_path}")

    def validate(self) -> bool:
        """Validate the configuration."""
        if not self.watch_folder:
            logger.error("Watch folder is not configured.")
            return False

        if not os.path.isdir(self.watch_folder):
            logger.error(f"Watch folder does not exist: {self.watch_folder}")
            return False

        if not self.printer_name:
            logger.warning("No printer configured. Will use default printer.")

        if self.print_method == "acrobat":
            if not self.acrobat_path or not os.path.isfile(self.acrobat_path):
                logger.error(f"Acrobat executable not found: {self.acrobat_path}")
                logger.info("Set 'acrobat_path' in config or use 'print_method': 'shellexecute'")
                return False

        return True

    def _find_acrobat_path(self) -> str:
        """Auto-detect Adobe Acrobat/Reader installation path."""
        possible_paths = [
            # Acrobat DC (paid version)
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            # Acrobat Reader DC (free version)
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            # Older versions
            r"C:\Program Files\Adobe\Reader 11.0\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
        ]

        for path in possible_paths:
            if os.path.isfile(path):
                logger.info(f"Auto-detected Acrobat at: {path}")
                return path

        return ""


def get_available_printers() -> list:
    """Get a list of available printers on the system."""
    printers = []
    try:
        for printer in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        ):
            printers.append(printer[2])
    except Exception as e:
        logger.error(f"Error getting printers: {e}")
    return printers


def get_default_printer() -> Optional[str]:
    """Get the default printer name."""
    try:
        return win32print.GetDefaultPrinter()
    except Exception as e:
        logger.error(f"Error getting default printer: {e}")
        return None


def print_pdf_acrobat(file_path: str, printer_name: str, acrobat_path: str) -> bool:
    """
    Print a PDF file using Adobe Acrobat/Reader command line.

    Uses the /t switch which prints to specified printer and exits.
    Command: Acrobat.exe /t "file.pdf" "printer_name"
    """
    try:
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return False

        if not os.path.isfile(acrobat_path):
            logger.error(f"Acrobat executable not found: {acrobat_path}")
            return False

        # Use the default printer if none specified
        if not printer_name:
            default = get_default_printer()
            if not default:
                logger.error("No printer available")
                return False
            printer_name = default

        logger.info(f"Printing with Acrobat: {file_path} to {printer_name}")

        # Build the command
        # /t = print to specified printer and exit
        # /h = start minimized
        cmd = [
            acrobat_path,
            "/t",
            file_path,
            printer_name
        ]

        logger.info(f"Executing: {' '.join(cmd)}")

        # Run Acrobat with the print command
        # Using subprocess.Popen so we don't block waiting for Acrobat
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        # Wait a bit for the print job to be submitted
        # Acrobat needs time to start, process the PDF, and send to printer
        time.sleep(5)

        logger.info(f"Print job sent successfully via Acrobat: {file_path}")
        return True

    except Exception as e:
        logger.error(f"Error printing {file_path} with Acrobat: {e}")
        return False


def print_pdf_shellexecute(file_path: str, printer_name: str) -> bool:
    """
    Print a PDF file using Windows ShellExecute.

    Uses the 'print' verb which works with the default PDF application.
    """
    try:
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return False

        # Use the default printer if none specified
        if not printer_name:
            default = get_default_printer()
            if not default:
                logger.error("No printer available")
                return False
            printer_name = default

        logger.info(f"Printing with ShellExecute: {file_path} to {printer_name}")

        # Set the printer as default temporarily for ShellExecute
        original_printer = get_default_printer()

        try:
            # Set our target printer as default
            win32print.SetDefaultPrinter(printer_name)

            # Use ShellExecute with 'print' verb
            # This uses the default PDF application's print functionality
            win32api.ShellExecute(
                0,  # Handle to parent window
                "print",  # Operation (print)
                file_path,  # File to print
                None,  # Parameters
                os.path.dirname(file_path),  # Working directory
                0,  # Show command (0 = hide)
            )

            logger.info(f"Print job sent successfully: {file_path}")
            return True

        finally:
            # Restore original default printer
            if original_printer and original_printer != printer_name:
                try:
                    win32print.SetDefaultPrinter(original_printer)
                except Exception:
                    pass  # Ignore errors restoring default printer

    except Exception as e:
        logger.error(f"Error printing {file_path}: {e}")
        return False


def print_pdf(file_path: str, printer_name: str, config: "Config") -> bool:
    """
    Print a PDF file using the configured method.

    Dispatches to either Acrobat command line or ShellExecute based on config.
    """
    if config.print_method == "acrobat":
        return print_pdf_acrobat(file_path, printer_name, config.acrobat_path)
    else:
        return print_pdf_shellexecute(file_path, printer_name)


def is_file_ready(file_path: str, timeout: int = 30) -> bool:
    """
    Check if a file is ready (fully written/synced) by attempting to open it exclusively.

    This is important for synced files that may still be being written.
    """
    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            # Try to open the file exclusively
            with open(file_path, "rb") as f:
                # Try to read a bit to ensure it's accessible
                f.read(1024)

            # Check if file size is stable
            size1 = os.path.getsize(file_path)
            time.sleep(0.5)
            size2 = os.path.getsize(file_path)

            if size1 == size2 and size1 > 0:
                return True

        except (IOError, OSError, PermissionError):
            pass

        time.sleep(0.5)

    return False


class PDFHandler(FileSystemEventHandler):
    """Handler for file system events in the watched folder."""

    def __init__(self, config: Config):
        self.config = config
        self.processed_files: set = set()

    def on_created(self, event: FileCreatedEvent) -> None:
        """Handle file creation events."""
        if event.is_directory:
            return

        file_path = event.src_path
        file_ext = os.path.splitext(file_path)[1].lower()

        # Check if it's a file type we care about
        if file_ext not in [ext.lower() for ext in self.config.file_extensions]:
            return

        # Avoid processing the same file twice
        if file_path in self.processed_files:
            return

        logger.info(f"New file detected: {file_path}")
        self.process_file(file_path)

    def process_file(self, file_path: str) -> None:
        """Process a new file for printing."""
        try:
            # Wait for file to be fully synced/written
            logger.info(f"Waiting for file to be ready: {file_path}")

            if not is_file_ready(file_path):
                logger.warning(f"File not ready after timeout: {file_path}")
                return

            # Additional delay to ensure sync is complete
            if self.config.print_delay_seconds > 0:
                time.sleep(self.config.print_delay_seconds)

            # Mark as processed to avoid duplicates
            self.processed_files.add(file_path)

            # Print the file
            success = print_pdf(file_path, self.config.printer_name, self.config)

            # Additional delay to ensure printing runs
            if self.config.move_delete_delay_seconds > 0:
                time.sleep(self.config.move_delete_delay_seconds)

            if success:
                self.handle_after_print(file_path)
            else:
                logger.error(f"Failed to print: {file_path}")
                # Remove from processed so it can be retried
                self.processed_files.discard(file_path)

        except Exception as e:
            logger.error(f"Error processing file {file_path}: {e}")
            self.processed_files.discard(file_path)

    def handle_after_print(self, file_path: str) -> None:
        """Handle file after successful printing (move or delete)."""
        try:
            if self.config.delete_after_print:
                os.remove(file_path)
                logger.info(f"Deleted after printing: {file_path}")

            elif self.config.move_after_print:
                # Create printed folder if it doesn't exist
                printed_dir = os.path.join(
                    self.config.watch_folder, self.config.printed_folder
                )
                os.makedirs(printed_dir, exist_ok=True)

                # Move file with timestamp to avoid conflicts
                filename = os.path.basename(file_path)
                name, ext = os.path.splitext(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = f"{name}_{timestamp}{ext}"
                new_path = os.path.join(printed_dir, new_filename)

                shutil.move(file_path, new_path)
                logger.info(f"Moved to: {new_path}")

        except Exception as e:
            logger.error(f"Error handling file after print: {e}")


def process_existing_files(config: Config, handler: PDFHandler) -> None:
    """Process any existing files in the watch folder on startup."""
    logger.info("Checking for existing files in watch folder...")

    watch_path = Path(config.watch_folder)

    for ext in config.file_extensions:
        for file_path in watch_path.glob(f"*{ext}"):
            if file_path.is_file():
                # Skip files in the printed folder
                if config.printed_folder in str(file_path):
                    continue

                logger.info(f"Found existing file: {file_path}")
                handler.process_file(str(file_path))


def list_printers() -> None:
    """Display available printers and exit."""
    print("\n" + "=" * 50)
    print("Available Printers")
    print("=" * 50)

    printers = get_available_printers()
    default_printer = get_default_printer()

    if not printers:
        print("No printers found.")
        return

    for printer in printers:
        if printer == default_printer:
            print(f"  * {printer} (default)")
        else:
            print(f"    {printer}")

    print("=" * 50)
    print("\nCopy the printer name to your config.json file.")
    print("Leave 'printer_name' empty to use the default printer.\n")


def main() -> None:
    """Main entry point."""
    print("\n" + "=" * 50)
    print("Folder2Print - Automatic PDF Printing")
    print("=" * 50 + "\n")

    # Check for --list-printers argument
    if len(sys.argv) > 1 and sys.argv[1] in ["--list-printers", "-l"]:
        list_printers()
        return

    # Load configuration
    config_path = "config.json"
    if len(sys.argv) > 1:
        config_path = sys.argv[1]

    try:
        config = Config(config_path)
    except SystemExit:
        return

    if not config.validate():
        logger.error("Configuration validation failed. Please check your config.json")
        list_printers()
        return

    # Show current configuration
    logger.info(f"Watch folder: {config.watch_folder}")
    logger.info(f"Printer: {config.printer_name or '(default)'}")
    logger.info(f"File types: {config.file_extensions}")

    # Create event handler and observer
    handler = PDFHandler(config)
    observer = Observer()
    observer.schedule(handler, config.watch_folder, recursive=False)

    # Process existing files first
    process_existing_files(config, handler)

    # Start watching
    observer.start()
    logger.info(f"Watching folder: {config.watch_folder}")
    logger.info("Press Ctrl+C to stop...")

    try:
        while True:
            time.sleep(config.check_interval_seconds)
    except KeyboardInterrupt:
        logger.info("Stopping...")
        observer.stop()

    observer.join()
    logger.info("Stopped.")


if __name__ == "__main__":
    main()
