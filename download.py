
import os
from SharePoint import Utils, SharePoint

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    #object for utils
    ut = Utils()
    #read config default from current path config.json
    config = ut.read_config_file()

    current_path = ROOT_DIR + '/ouput.csv' or "/tmp/output.csv"

    #set the download_args
    try:
        download_args = {
            "local_dest_file_path": current_path,
            "site_doc":  config['doc_library'] if  config['doc_library'] is not None else "",
            "sharepoint_src": config['source_file_name'] if config['source_file_name'] is not None else "",
            "sharepoint_src_name":config['source_file_name'] if config['source_file_name'] is not None else "",
        }
    except Exception as e:
        print("Config key not found ", e)
        return False

    #get user name passwor
    username, password = ut.get_username_password()
    #create share point object
    sp = SharePoint(username, password )
    #call download file
    result = sp.download_file(**download_args)
    print(result)

#run the program
if __name__ == "__main__":
    main()