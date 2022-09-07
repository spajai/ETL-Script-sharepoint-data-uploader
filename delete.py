
import os
from SharePoint import Utils, SharePoint

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    #object for utils
    ut = Utils()
    #read config default from current path config.json
    config = ut.read_config_file()

    #set the upload_args
    try :
        delete_args = {
            "sharepoint_doc":  config['sharepoint_doc'] if  config['sharepoint_doc'] is not None else "",
            "sharepoint_url":  config['sharepoint_url'] if  config['sharepoint_url'] is not None else "",
            "sharepoint_src": config['sharepoint_src'] if config['sharepoint_src'] is not None else "",
            "delete_sharepoint_file_name": config['delete_sharepoint_file_name'] if config['delete_sharepoint_file_name'] is not None else "",
        }
    except Exception as e:
        print("Config key not found ", e)
        return False

    #get username password
    username, password = ut.get_username_password()
    #create share point object
    sp = SharePoint(username, password )
    #call upload file
    result = sp.delete_file(**delete_args)
    print(result)

#run the program
if __name__ == "__main__":
    main()