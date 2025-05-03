import sys
import os
import subprocess
import platform
from pathlib import Path

from classes.control import ctrlCommon
from classes.control import ctrlMessage


class ctrlBatch():
    # Get the file path
    def get_path(relative_path):
        if hasattr(sys, '_MEIPASS'):
            # Path when running as an EXE
            return os.path.join(sys._MEIPASS, relative_path)
        else:
            # Path during debugging
            return os.path.join(
                os.path.abspath("."),
                'classes/script',
                relative_path)
    
    def _test_pwsh():
        # Pre-check: Verify that pwsh can run properly and that the version is 7 or higher.
        try:
            # Use $PSVersionTable.PSVersion.Major to get the major version
            version_proc = subprocess.run(
                ["pwsh", "-NoProfile", "-Command", "$PSVersionTable.PSVersion.Major"],
                capture_output=True, text=True
            )
            
            # If the command fails
            if version_proc.returncode != 0:
                ctrlMessage.print_error("Failed to check the PowerShell version. Please ensure pwsh is working correctly.")
                return -4001
            
            major_version_str = version_proc.stdout.strip()

            _result = 0
            try:
                major_version = int(major_version_str)
            except ValueError:
                ctrlMessage.print_error(f"An exceptional error occurred while retrieving the PowerShell version (non-numeric version information): {major_version_str}")
                _result = -4101
            
            if major_version < 7:
                ctrlMessage.print_error(f"The PowerShell version is less than 7 (actual version: {major_version}). PowerShell 7 or later is required.")
                _result = -4102
        
        except FileNotFoundError:
            ctrlMessage.print_error("The pwsh command was not found. PowerShell 7 or later may not be installed, or the system may need to be restarted after installation. Alternatively, check the environment variable definitions.")
            _result = -4111
        except Exception as err:
            ctrlMessage.print_error(f"An error occurred while checking the pwsh version: {err}")
            _result = -4112
            
        return _result

    def exe_powershell(script_path, *args):
        _result = 0

        # Test pwsh execution
        _result = ctrlBatch._test_pwsh()

        # Prepare arguments for running the PowerShell script
        if _result == 0:
            pscommand = [
                'pwsh',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-file',
                script_path
            ]
            # pscommand = [
            #     'powershell',
            #     '-NoProfile',
            #     '-ExecutionPolicy',
            #     'Bypass',
            #     '-file',
            #     script_path
            # ]

            # Add arguments for the PowerShell script if any
            if args:
                pscommand.extend(args)

            # Abort processing if the OS is not Windows
            osname = platform.system()
            if osname != 'Windows':
                _result = -4201
            
            if _result == 0:
                try:
                    CREATE_NO_WINDOW = 0x08000000
                    _compps = subprocess.run(pscommand, creationflags=subprocess.CREATE_NO_WINDOW)
                    _result = _compps.returncode
                    # If the return value is a positive integer, convert it to a negative value
                    if _result > 2**31 - 1:
                        _result -= 2**32
                except Exception as err:
                    _result = -4202
                    ctrlMessage.print_error(err)
        
        return _result
    
    def open_folder(target_path):
        _result = 0

        _open_path = Path(target_path)
        if not _open_path.exists():
            _result = -4301
        elif sys.platform.startswith("win"):
            try:
                subprocess.Popen(["explorer", str(_open_path)])
            except Exception as err:
                _result = -4302
                ctrlMessage.print_error(err)
        
        return _result