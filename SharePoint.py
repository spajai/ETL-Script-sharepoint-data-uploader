import json, os,sys, glob,csv, re
import pandas as pd

from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

class Utils:
    # default constructor
    def __init__(self):
        pass

    def get_local_xlsb_files(self, local_dir) :
        """get_local_xlsb_files
        self :invocant
        :param str local_dir: 
        return: valid xlsb files directory
        """
        all_file_path = None
        xlsb_files = glob.glob(local_dir + '*.xlsb')
        if len(xlsb_files) < 1 :
            print("[ERROR] exiting No .xlsb Files found to process in ",local_dir)
            sys.exit(2)

        #iterate on all files
        for file_path in xlsb_files:
            if not os.path.isfile(file_path) and os.stat(file_path).st_size  < 1:
                print("Inavlid File path ", file_path)
                sys.exit(2)
            else :
                all_file_path.append(file_path)
        return all_file_path

    def read_config_file (self, file_path = None) :
        """read_config_file
        self :invocant
        str file_path : json file path 
        return: config
        """
        if file_path is None :
            ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(ROOT_DIR, 'config.json')

        try:
            # read config file
            with open(file_path) as config_file:
                config = json.load(config_file)
                config = config['share_point']
            self.config = config
            return config
        except:
            print("[ERROR] Invalid conf or Conf file path ", file_path)
            sys.exit(2)
        else:
            print("Valid conf and env variable found")


 
    def delete_local_file (self, local_file_path):
        """delete_local_file
        self :invocant
        str input_file_path : input xlsb file path eg '/tmp/input.xlsb'
        str output_file_path : output csv file path  eg '/tmp/out.csv'
        return: true or false
        """

        if os.path.isfile(local_file_path) :
            os.remove(local_file_path)
        else :
            print("[ERROR] invalid path or not a file ", local_file_path)

    def get_username_password (self) :
        """get_username_password
        self :invocant
        This method will set the username and password from environment varible USER and PASSWRORD
        if not found will check the config but read_config_file has to be called before calling this method
        if both the places failed to return the username password then it will fail

        return: username, password
        """
        username , password = None, None
        try:
            username = os.environ['USER']
            password = os.environ['PASSWORD']
        except Exception as e:
             print("environment varible missing " + format(e) + " checking config file")
             if bool(self.config) :
                username = self.config['user']
                password = self.config['password']
        if not (bool(username) and bool(password)) :
            print("[ERROR] unable to set the username and password from environment or conf file")
            print("[HINT] Set environment varible USER , PASSWORD  or read conf bnefore calling this def")
            sys.exit(2)
        else:
            return username, password


class  SharePoint():
    # default constructor
    def __init__(self, usr = None, passwd = None):

        #default var
        username  = None
        password = None

        if os.environ.get('USERNAME') is not None:
            username = os.environ.get('USERNAME')
        else :
            username = usr 

        if os.environ.get('PASSWORD') is not None:
            password = os.environ.get('PASSWORD')
        else :
            password = passwd

        if  bool(username) and bool(password) :
            self.credential_store = UserCredential(username, password)
        else :
            print("Environment varible USERNAME,PASSWORD not set or object creation missed username,password")
            sys.exit(2)

    def delete_file (self, **delete_args):
        """delete the file from sharpoint
        self invocant
        :param str sharepoint_doc:
        :param str sharepoint_src:
        :param str sharepoint_url:
        :param str sharepoint_file_name:

        return: none
        """
        relative_url = "{doc}{src_dir}{src_file_name}".format(doc=delete_args["sharepoint_doc"],src_dir=delete_args["sharepoint_src"],src_file_name=delete_args["sharepoint_file_name"])
        ctx = ClientContext(delete_args["sharepoint_url"]).with_credentials(self.credential_store)
        file_to_delete = ctx.web.get_file_by_server_relative_url(relative_url)  
        file_to_delete.delete_object()
        ctx.execute_query()

    def upload_file(self, **upload_args):
        """upload the file to sharpoint
        :param str sharepoint_url:
        :param int cloud_path:
        :param str local_file_path:

        return: True / False
        """
        try :
            ctx = ClientContext(upload_args["sharepoint_url"]).with_credentials(self.credential_store)
            target_folder = ctx.web.get_folder_by_server_relative_url(upload_args["cloud_path"])
            size_chunk = 1000000
            file_size = os.path.getsize(upload_args["local_file_path"])

            # def print_upload_progress (offset,file_size):
            #     print("Uploaded '{0}' bytes from '{1}'...[{2}%]".format(offset, file_size, round(offset / file_size * 100, 2)))

            print_upload_progress = lambda offset,file_size : (
               sys.stdout.write("Uploaded '{0}' bytes from '{1}'...[{2}%]".format(offset, file_size, round(offset / file_size * 100, 2)))
            )

            uploaded_file = target_folder.files.create_upload_session(upload_args["local_file_path"], size_chunk, print_upload_progress(size_chunk,file_size)).execute_query
            print('File {0} has been uploaded successfully'.format(uploaded_file.serverRelativeUrl))
            return True
        except Exception as e:
            print(e)
            return False

    def download_file (self, **download_args):
        """deletes the file from sharpoint
        :param str local_dest_file_path: 
        :param str site_doc: 
        :param str sharepoint_src:
        :param str sharepoint_src_name:

        return: none
        """
        local_file_path = download_args["local_dest_file_path"]
        with open(local_file_path , 'wb') as local_file:
            try :
                abs_file_url = "{site_doc}{src_dir}{src_file_name}".format(site_doc=download_args["site_doc"],src_dir=download_args["sharepoint_src"],src_file_name=download_args["sharepoint_src_name"])
                file = File.from_url(abs_file_url).with_credentials(self.credential_store).download(local_file).execute_query()
                print("'{0}' file has been downloaded into {1}".format(file.serverRelativeUrl, local_file.name))
                return True
            except Exception as e:
                if os.stat(local_file_path).st_size  < 1:
                    print("[ERROR] Invalid file size due to download error deleting", local_file_path)
                    os.remove(local_file_path)
                    sys.exit(2)

class TransformCSV:
        # default constructor
    def __init__(self):
        pass

    def transform_xlsb_to_csv (self, input_file_path, output_file_path) :
        """transform_xlsb_to_csv
        self :invocant
        str input_file_path : input xlsb file path eg '/tmp/input.xlsb'
        str output_file_path : output csv file path  eg '/tmp/out.csv'
        return: true or false
        """
        print("Processing Input ==> ", input_file_path)

        try:
            if os.path.isfile(input_file_path) and os.stat(input_file_path).st_size  > 1 :
                pass
        except:
            print("Inavlid File path ", input_file_path)
            sys.exit(2)

        try:
            #read file
            df = pd.read_excel(input_file_path, engine='pyxlsb',  index_col=None, header=None,na_filter=0,skipfooter=0)
            result_data = {"header": [],"header_values": [],"out_data": [],"quarter_header": {}}

            #transpose
            for i, row in df.iterrows():
            #consider the row till here as columns
                if i <= 12:
                    result_data["header"].append(row[0])
                    result_data["header_values"].append(row[1])
                elif i > 12 and i < 15:
                    continue
                elif i == 15:
                    for j in range (15):
                        if row[j] != '':
                            result_data["quarter_header"][row[j]] = 1
                elif i == 16:
                    for j in range (6):
                        result_data["header"].append(row[j])
                else:
                    out_row = []
                    #colum length zero last empty rows added by spreadsheet
                    if len((row[0] + row[1] + row[2]).strip()) == 0:
                        continue
                    #remove last row for total
                    elif re.match( r'.*total.*', row[0]+row[1], re.M|re.I):
                        continue
                    for j in range (15):
                        # append the data
                        out_row.append(row[j])
                    result_data["out_data"].append(out_row)

            #get quater keys
            qtr = list(result_data["quarter_header"].keys())
            result = []
            for data in result_data["out_data"]:
                col = 4
                for qyr in qtr:
                    row = list()
                    q_yr  = qyr[0:4]
                    q_num = qyr[4:]
                    #expand to Net | Marg |Marg %
                    row.extend(result_data["header_values"])
                    #add 3 column data
                    row.extend(data[0:4])
                    # add quater yr and num
                    row.extend([q_yr,q_num])
                    # append quater values replace empty with 0
                    q_data_wi_zero = [i if i != '' else 0 for i in data[col:col+3]]
                    row.extend(q_data_wi_zero)
                    col = col + 3
                    result.append(row)
            #append custom header
            result_data["header"].insert(17,'Ryear')
            result_data["header"].insert(18,'RQuater')
            for i in range(len(result)):
                for j in range(-3, 0):
                    result[i][j] = "{:.15f}".format(float(result[i][j]))
            #open output file
            if output_file_path is None :
                output_file_path = os.path.splitext(input_file_path)[0] + '.csv'

            with open(output_file_path, 'w', newline='\n') as file:
                writer = csv.writer(file)
                writer.writerows([result_data["header"]])
                writer.writerows(result)
                print("Writing output to file => ",output_file_path)
                return True
        except Exception as e:
            print("[ERROR] ", e)
            return False

class All (SharePoint,Utils, TransformCSV):
    pass