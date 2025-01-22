import os
import shutil


def delete_all_in_directory(directory_path):
    # Loop over all the items in the directory
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)
        if os.path.isdir(item_path):
            # If it's a directory, delete it
            shutil.rmtree(item_path)
        else:
            # If it's a file, delete it
            os.remove(item_path)