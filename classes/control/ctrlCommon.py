import sys
import os


class ctrlCommon():
    # Get the file path
    def get_path(relative_path, filename):
        if hasattr(sys, '_MEIPASS'):
            # Path when executing as an EXE
            return os.path.join(sys._MEIPASS, relative_path, filename)
        else:
            # Path during debugging
            return os.path.join(
                os.path.abspath("."),
                relative_path,
                filename)
