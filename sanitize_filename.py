# Helper function to sanitize folder and file names
import re


def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*\t\n\r]', '_', str(name))
    # Read the input file
