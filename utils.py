import os

def get_filenames_in(path, file_extension=None):
    filenames = os.listdir(path)
    if file_extension:
        filenames = [filename for filename in filenames if filename.endswith(file_extension)]
    return filenames
    
