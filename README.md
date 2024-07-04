# CDDpassremove
This repository contains a Python script designed to batch remove passwords from Microsoft Excel '.xls' files. The script uses the 'win32com.client' library to interact with Excel and remove the password protection from each file in a specified directory.

## Libraries
pip install pywin32

## USAGE
python CDDpasswd.py <directory_path> <password_to_remove>

## Actual need of this tool
When we have numerous password-protected .xls files that need to be sent to client by removing the passwords, manually doing this becomes time-consuming.  
