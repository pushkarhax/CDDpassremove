import os
import win32com.client

def remove_password(file_path, password):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Open the workbook with the provided password
        workbook = excel.Workbooks.Open(file_path, False, False, None, password)
        # Unprotect the workbook
        workbook.Unprotect(password)
        
        # Save the workbook without password
        workbook.SaveAs(
            file_path, 
            FileFormat=56,  # 56 is for .xls format
            Password='',
            WriteResPassword='',
            ReadOnlyRecommended=False,
            CreateBackup=False
        )
        
        workbook.Close(SaveChanges=True)
        print("Password removed for: {}".format(file_path))
    except Exception as e:
        print("Failed to remove password for {}: {}".format(file_path, e))
    finally:
        excel.Quit()

def process_files(directory, password):
    for filename in os.listdir(directory):
        if filename.endswith(".xls"):
            file_path = os.path.join(directory, filename)
            remove_password(file_path, password)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='Remove password protection from Excel files.')
    parser.add_argument('directory', type=str, help='Directory containing Excel files')
    parser.add_argument('password', type=str, help='Password to remove from Excel files')

    args = parser.parse_args()
    process_files(args.directory, args.password)
import os
import win32com.client

def remove_password(file_path, password):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Open the workbook with the provided password
        workbook = excel.Workbooks.Open(file_path, False, False, None, password)
        # Unprotect the workbook
        workbook.Unprotect(password)
        
        # Save the workbook without password
        workbook.SaveAs(
            file_path, 
            FileFormat=56,  # 56 is for .xls format
            Password='',
            WriteResPassword='',
            ReadOnlyRecommended=False,
            CreateBackup=False
        )
        
        workbook.Close(SaveChanges=True)
        print("Password removed for: {}".format(file_path))
    except Exception as e:
        print("Failed to remove password for {}: {}".format(file_path, e))
    finally:
        excel.Quit()

def process_files(directory, password):
    for filename in os.listdir(directory):
        if filename.endswith(".xls"):
            file_path = os.path.join(directory, filename)
            remove_password(file_path, password)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='Remove password protection from Excel files.')
    parser.add_argument('directory', type=str, help='Directory containing Excel files')
    parser.add_argument('password', type=str, help='Password to remove from Excel files')

    args = parser.parse_args()
    process_files(args.directory, args.password)
