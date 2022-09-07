import os
from SharePoint import TransformCSV

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

def main():

    #set the I/O file path
    input_file_path = ROOT_DIR + '/input.xlsb' or "/tmp/Project.xlsb"
    output_file_path = ROOT_DIR + '/ouput.csv' or "/tmp/output.csv"

    #create TransformCSV obj
    result = TransformCSV().transform_xlsb_to_csv(input_file_path, output_file_path)
    print(result)

#run the program
if __name__ == "__main__":
    main()