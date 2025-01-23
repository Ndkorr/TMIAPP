import subprocess



def install_library(library_name):

    """
    Installs a Python library using pip.
    
    Args:
        library_name (str): The name of the Python library to install.
    """

    command = ["pip", "install", library_name]

    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    output, error = process.communicate()

    

    if process.returncode == 0:

        print(f"Successfully installed {library_name}.")

    else:

        print(f"Error installing {library_name}: {error.decode('utf-8')}")



# Example usage:

install_library("pip") 

