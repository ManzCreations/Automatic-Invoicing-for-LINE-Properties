import os
import shutil


def remove_directory_if_exists(directory):
    """
    Remove a directory and its contents if it exists.
    """
    if os.path.exists(directory):
        shutil.rmtree(directory, ignore_errors=True)
        print("Directory removed successfully!")
    else:
        print("Directory does not exist.")


def setup_directory_structure(app, month_name, year):
    """
    Setup and clean directory structure for reports.
    """
    cwd = os.getcwd()
    path_year = os.path.join(cwd, f'{year} Reports')
    path_month = os.path.join(path_year, f'{month_name} Report')

    remove_directory_if_exists(path_month)

    if not os.path.exists(path_month):
        os.makedirs(path_month)
        app.log('The new directory is created!')

    return path_month
