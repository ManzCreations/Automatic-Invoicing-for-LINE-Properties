from data_utilities import initialize_dataframes


def load_files(app, filepath, filenames, path):
    """
    Load necessary files and create dataframes.

    Args:
        app: Application context for logging.
        filepath (str): Path to the model files.
        filenames (list): List of filenames to process.
        path (str): Path to the report directory.
    """
    reformat_info = {
        'bnb': {'path': '', 'sheet': ''},
        'check': {'path': '', 'sheet': ''},
        'vrbo': {'path': '', 'sheet': ''}
    }

    num_files = len(os.listdir(filepath))
    progress_vals = np.linspace(3, 15, num=num_files)
    file_idx = 0

    dataframes = initialize_dataframes()

    for fil in os.listdir(filepath):
        app.progress_bar["value"] = int(progress_vals[file_idx])
        app.progress_bar.update()
        file_idx += 1
        file_full_path = os.path.join(filepath, fil)

        for fname in range(len(filenames)):
            if fil.lower().startswith(filenames[fname].lower()):
                df = pd.read_excel(file_full_path) if fil.endswith('.xlsx') else pd.read_csv(file_full_path)
                reformat_info_key = list(reformat_info.keys())[fname] if fname < len(reformat_info) else None

                if reformat_info_key:
                    reformat_info[reformat_info_key]['path'] = os.path.join(path, fil)
                    reformat_info[reformat_info_key]['sheet'] = fil

                if fname == 2:  # Special handling for Current with multiple sheets
                    xl = pd.ExcelFile(file_full_path)
                    for des_sub_name in ['Cleaning', 'Customer']:
                        sheet = [s for s in xl.sheet_names if des_sub_name in s]
                        if sheet:
                            name_str = ''.join(sheet)
                            dataframes[des_sub_name.lower()] = pd.read_excel(file_full_path, sheet_name=name_str)
                            if des_sub_name == 'Customer':
                                copyfile(file_full_path, os.path.join(path, fil))
                else:
                    dataframes[list(dataframes.keys())[fname]] = df

    return dataframes, reformat_info
