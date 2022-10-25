import pandas as pd
import json
import os


# File system
def remove_file(target_file: str):
    """
    Remove the target file
    """
    if os.path.exists(target_file):
        os.remove(target_file)

def is_allowed_file(file_name: str, allowed_extensions: list):
    """
    Method For checking exetension
    """
    allowed_extensions = set(allowed_extensions)
    return "." in file_name and file_name.rsplit(".", 1)[1] in allowed_extensions

def is_excel_file(file_name: str):
    """
    To check if it is excel file.
    """
    return is_allowed_file(file_name, ["xlsx", "xls"])

def traversing_dir_for_file(target_path: str, target_condition):
    """
    Get file list that are target exetension below the target path. (recursive)
    """
    matched_files = []
    for root, _, files in os.walk(target_path):
        for file in files:
            file_full_path = os.path.join(root, file)
            if target_condition(file_full_path):
                matched_files.append(file_full_path)
    return matched_files

# Merge Logic

def parse_config(file_name:str="./config.json"):
    """
    Parsing the config.json file
    """
    with open(file_name) as f:
        json_data = json.load(f)
        return json_data

def load_excel_file(file_name:str, sheet_num:int=0) -> pd.DataFrame:
    """
    Load a excel file
    """
    df = pd.read_excel(file_name)
    return df


def main():
    """
    The main function
    """
    # For config information
    config_info = parse_config()
    root_path = config_info["root_path"]
    step_col_name = config_info["step_col_name"]
    target_step_num = config_info["target_step_num"]
    result_categories = config_info["result_categories"]

    # Delete old merge_data.xlsx
    result_file_path = ".".join(["merge_data",'xlsx'])
    remove_file(result_file_path)

    target_files = sorted(traversing_dir_for_file(root_path, is_excel_file))
    for target_file in target_files:
        df = load_excel_file(target_file)
        target_df = df[df[step_col_name]==target_step_num]
        target_df["date_file"] = "_".join([os.path.dirname(target_file).split('\\')[-1], os.path.basename(target_file).split('.')[0]])
        
        for category_name, col_names in result_categories.items():
            category_data = target_df[['date_file', *col_names]]

            (mode, if_sheet_exists) = ('a', 'overlay') if os.path.exists(result_file_path) else ('w', None)
            with pd.ExcelWriter(result_file_path, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer: # pylint: disable=abstract-class-instantiated
                if category_name in writer.sheets:
                    category_data.to_excel(writer, sheet_name=category_name, startrow=writer.sheets[category_name].max_row, header=False, index=False)
                else:
                    category_data.to_excel(writer, sheet_name=category_name, index=False)

if __name__=="__main__":
    main()