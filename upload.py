
import os
from SharePoint import Utils, SharePoint

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    #object for utils
    ut = Utils()
    #read config default from current path config.json
    config = ut.read_config_file()

    local_file_path = ROOT_DIR + '/ouput.csv' or "/tmp/output.csv"

    #set the upload_args
    try :
        upload_args = {
            "local_file_path": local_file_path,
            "sharepoint_url":  config['upload_sharepoint_url'] if  config['upload_sharepoint_url'] is not None else "",
            "cloud_path": config['upload_cloud_path'] if config['upload_cloud_path'] is not None else "",
        }
    except Exception as e:
        print("Config key not found ", e)
        return False

    #get username password
    username, password = ut.get_username_password()
    #create share point object
    sp = SharePoint(username, password )
    #call upload file
    result = sp.upload_file(**upload_args)
    print(result)

#run the program
if __name__ == "__main__":
    main()