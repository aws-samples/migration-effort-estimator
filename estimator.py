import json
import zipfile
from datetime import datetime
import time, argparse, os, shutil, ntpath, logging, boto3, random, openpyxl
from pathlib import Path
import pandas as pd
from io import StringIO # python3; python2: BytesIO 

logger = logging.getLogger()
logger.setLevel(logging.NOTSET)

# console handler to log errors
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.ERROR)
console_handler_format = '%(asctime)s | %(levelname)s: %(message)s'
console_handler.setFormatter(logging.Formatter(console_handler_format))
logger.addHandler(console_handler)

#file handler for detailed logging
file_name='Logs/hadoop-estimator-{:%Y-%m-%d-%H-%m-%S}.log'.format(datetime.now())
file_handler = logging.FileHandler('Logs/hadoop-estimator-{:%Y-%m-%d-%H-%m-%S}.log'.format(datetime.now()))
file_handler.setLevel(logging.INFO)
file_handler_format = '%(asctime)s | %(levelname)s | %(lineno)d: %(message)s'
file_handler.setFormatter(logging.Formatter(file_handler_format))
logger.addHandler(file_handler)

def write_to_estimator_template(res,output):
    try: 
        output=output.split('/')[1].replace('.zip','')
        #print(output)
        files = json.loads(res)
        TemplateFilePath="input/estimator_template.xlsx"
        SummaryPath="Excels/"+output+"_estimator_summary.xlsx"
        filePath="Excels/"+output+"_estimator_file_level.xlsx"
        xfile = openpyxl.load_workbook(TemplateFilePath)
        xfile.save(SummaryPath)
        xfile = openpyxl.load_workbook(TemplateFilePath)
        xfile.save(filePath)
        #shutil.copy(TemplateFilePath,SummaryPath)
        #shutil.copy(TemplateFilePath,filePath)
        files_list = []
        json_data=[]
        for iterator in files:
            for file in files[iterator]:
                entry = files[iterator][file]
                keys = entry.keys()
                data ={}
                data['Feed Type'] = iterator
                data['File Name'] = file
                for key in keys:
                    data[key]=entry[key] 
                files_list.append(data)
        final_df = pd.DataFrame.from_dict(files_list)
        logger.info("\n Final Data frame : \n%s" %final_df)
        with pd.ExcelWriter(filePath,mode='a',if_sheet_exists='replace' ) as writer:  
            final_df.to_excel(writer, sheet_name='File_Level_Info')
        feed_type = final_df["Feed Type"].unique()
        xfile = openpyxl.load_workbook(filePath)
        yfile= openpyxl.load_workbook(SummaryPath)
        for row in yfile['Estimation Matrix']['D4:Q10']:
            for cell in row:
                cell.value = None
        config_sheet = xfile['Config']
        simple_hours = config_sheet['B2'].value
        simple_medium_hours = config_sheet['B3'].value
        medium_hours = config_sheet['B4'].value
        medium_complex_hours = config_sheet['B5'].value
        complex_hours = config_sheet['B6'].value
        sheet = yfile['Estimation Matrix']
        id = 4
        for feed in feed_type:
            sheet['D'+str(id)] = feed
            row_df = final_df[final_df['Feed Type']==feed]
            total_scripts = len(row_df)
            complex_count=(row_df['complexity']=='complex').sum()
            medium_count=(row_df['complexity']=='medium').sum()
            medium_complex_count=(row_df['complexity']=='medium-complex').sum()
            simple_count=(row_df['complexity']=='simple').sum()
            simple_medium_count=(row_df['complexity']=='simple-medium').sum()
            logger.info(" Feed Type : {} , Total Scripts : {} , simple : {} , 'simple-medium' : {}, Medium :{} , Medium-complex: {}, complex :{}".format(feed,total_scripts,simple_count,simple_medium_count,medium_count,medium_complex_count,complex_count))
            sheet['E'+str(id)] = total_scripts
            sheet['F'+str(id)] = "Re-Platform"
            sheet['G'+str(id)] = simple_count
            sheet['H'+str(id)] = simple_medium_count
            sheet['I'+str(id)] = medium_count
            sheet['j'+str(id)] = medium_complex_count
            sheet['k'+str(id)] = complex_count
            sheet['L'+str(id)] = simple_count*simple_hours
            sheet['m'+str(id)] = simple_medium_count*simple_medium_hours
            sheet['n'+str(id)] = medium_count*medium_hours
            sheet['o'+str(id)] = medium_complex_count* medium_complex_hours
            sheet['p'+str(id)] = complex_count* complex_hours
            sheet['q'+str(id)] = sheet['L'+str(id)].value+ sheet['m'+str(id)].value + sheet['n'+str(id)].value + sheet['o'+str(id)].value + sheet['p'+str(id)].value
            id = id+1
            yfile.save(SummaryPath)
        del xfile['Estimation Matrix']
        del xfile['Config']
        xfile.save(filePath)
        del yfile['File_Level_Info']
        yfile.save(SummaryPath)
        try:
            os.remove('Excels/.DS_Store')
        except Exception as exception:
            logger.info('No hidden files')
        return SummaryPath, filePath
    except Exception as exception:
        logger.error("Exception occured while writing to template")
        logger.error(exception)
        exit()
    
def upload_to_s3(source_bucket, SummaryPath, fileinfoPath):
    appname = 'Hadoop'
    source_bucket = source_bucket
    output_dir="Excels"
    message = ""
    message += "\n Migration Effort Estimator summary" + "\n\n"
    message += "##########################################################\n"
    try:
        for FilePath in [SummaryPath, fileinfoPath]:
            read_file = pd.read_excel (FilePath)
            csv_buffer = StringIO()
            read_file.to_csv(csv_buffer)
            target_key="ESTIMATOR/Output/{}/".format(appname)+FilePath.replace('Excels/','').split('.')[0]+".csv"
            s3_resource = boto3.resource('s3')
            s3_resource.Object(source_bucket, target_key).put(Body=csv_buffer.getvalue())
            message += "# App Name:- " + str(appname) + "\n"
            message += "# S3 Folder of Results:- " +source_bucket +'/'+ target_key + "\n"
        """
        for f in os.listdir(output_dir):
            print(f)
            FilePath = os.path.join(output_dir, f)
            print(FilePath)
            read_file = pd.read_excel (FilePath)
            csv_buffer = StringIO()
            read_file.to_csv(csv_buffer)
            target_key="ESTIMATOR/Output/{}/".format(appname)+f.split('.')[0]+".csv"
            s3_resource = boto3.resource('s3')
            s3_resource.Object(source_bucket, target_key).put(Body=csv_buffer.getvalue())
        """
        message += "##########################################################\n"
        return message
    
    except Exception as exception:
        logger.error("Exception occured while uploading to S3")
        logger.error(exception)
        exit()

def send_notification(body,appname, topic):
    # Create an SNS client
    try:
        sns = boto3.client('sns')
        # Publish a simple message to the specified SNS topic
        response = sns.publish(
            TargetArn=topic,  
            Message=body, 
            MessageAttributes= {
                'appname': {
                    'DataType': 'String',
                    'StringValue': appname.lower()
                    }
                    },
            Subject="Migration Effort Estimator Calculator Process is Completed at:{} ".format(datetime.now())
        )
    except Exception as exception:
        logger.error("Exception occured while Sending notification. Check your Topic ARN and AWS Configure for default region ?")
        logger.error(exception)

def unzip_file(input_zip, output_dir):
    """Unzipping file to output_dir; returns file path to output directory"""
    try:
        logger.info("Unzipping file: %s" %input_zip)
        with zipfile.ZipFile(input_zip, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
    except Exception as exception:
        logger.error("Exception occured while unzipping file")
        logger.error(exception)

    output_dir=Path.joinpath(Path(output_dir),(ntpath.split(input_zip)[1]).split(".")[0])
    logger.info("Extracted file to: %s" %output_dir)
    folder_structure(output_dir,output_dir)
    logger.info("Unzipping and sorting file to %s done" %output_dir)
    return output_dir

def folder_structure(output_dir, final_path):
    """Create folder structure and sort file"""
    try:
        
        for f in os.listdir(output_dir):
            """Path to the original file"""
            original_file_path = os.path.join(output_dir, f)
            logger.info("File %s" %original_file_path)
            if (f == '.DS_Store'):
                original_file_path = os.path.join(output_dir, '.DS_Store')
                #print(original_file_path)
                os.remove(original_file_path)
                continue
            """Only operate on files"""
            if os.path.isfile(original_file_path):
                """Get file name portion only"""
                
                file_name = os.path.basename(original_file_path)
                
                """Get the extension of the file and create a path for it"""
                extension = f.split(".")[-1]
                extension_path = os.path.join(final_path, extension)
                """Create the path for files with the extension if it doesn't exist"""
                if not os.path.exists(extension_path):
                    os.makedirs(extension_path)
                """Copy the files into the new directory """
                shutil.move(original_file_path, os.path.join(extension_path, file_name))
            else:
                logger.info("Nested folder %s" %original_file_path)
                folder_structure(original_file_path, final_path)
        return
    except Exception as exception:
        logger.error("Exception while creating folder structure")
        logger.error(exception)
        exit()

def complexity_of_file(res, config):
    """Finding complexity of file; returns a dict"""
    """Open App config file """
    try :
        file=open(config)
        """convert to lower case for string manipulation"""
        appconfig={k.lower():v for k,v in json.load(file).items()}
        parameters=[i.lower() for i in appconfig["keywords"]]
        """Add parameter  line_of_code to parameters to define complexity"""
        parameters.append("line_of_code")
        """Define levels of complexity"""
        levels=["complex", "medium-complex", "medium","simple-medium"]
        for level in levels:
            for parameter in parameters:
                """Comparing complexity parameter count in file and config file
                set the complexity level if atleast one parameter count > count defined in app config """
                try:
                    if res[parameter] > int(appconfig[parameter][level]):
                        res["complexity"]=level
                        return res
                except Exception as exception:
                    continue

        """Set complexity level to simple if none of the condition match """
        res["complexity"]="simple"
        return res
    except Exception as exception:
        logger.error("Exception while Calculating complexity")
        logger.error(exception)
        exit()

def file_length(file_path):
    try:
        with open(file_path, 'r', errors='ignore') as fp: 
            return len(fp.readlines())
    except Exception as exception:
        logger.error("Exception in file_length()")
        logger.error(exception)

def search_string_in_file(file_name, string_to_search):
    """Search for the given string in file and returns number of occurence of the same"""
    try:
        line_number = 0
        list_of_results = []
        # Open the file in read only mode
        with open(file_name, 'r', errors='ignore') as read_obj:
            # Read all lines in the file one by one
            for line in read_obj:
                # For each line, check if line contains the string
                line_number += 1
                if line.startswith('/*') or line.startswith('#'):
                    continue
                elif string_to_search.upper() in line.upper():
                    # If yes, then add the line number & line as a tuple in the list
                    list_of_results.append((line_number, line.rstrip()))
        # Return list of tuples containing line numbers and lines where string is found
        return len(list_of_results)

    except Exception as exception:
        logger.error("Exception in search_string_in_file")
        logger.error(file_name)
        logger.error(exception)
        exit()

def estimator(output_dir,config):
    """Search each keyword defined in app config in files 
    returns a dictionary with count of defined config parameter and complexity of file""" 
    try: 
        file=open(config)
        # convert to lower case for string manipulation
        data={k.lower():v for k,v in json.load(file).items()}
        keywords=[i.lower() for i in data["keywords"]]
        res_dict={}
        # For each extension flder in output directory
        for folder in os.listdir(output_dir):
                folder_dict={}
                # Constructing path for folder
                folder_path = os.path.join(output_dir, folder)
                # For each file in the folder
                for file in os.listdir(folder_path):
                    res={}
                    # Constructing path for file
                    file_path=os.path.join(folder_path, file)
                    """Find the count of each keyword defined in the app config file
                    and number of line in the file"""
                    for keyword in keywords:
                        res[keyword]=search_string_in_file(file_path,keyword)
                    # Add line count of file to result
                    res["line_of_code"]= file_length(file_path)
                    # Find the complexity of the file based on keywords count
                    res=complexity_of_file(res, config)
                    """append result of each file to a dictionary that maintains 
                    folder level details"""
                    folder_dict[file]=res
                # append result of folder to a dictionary to form result 
                res_dict[folder]=folder_dict
        logger.info("\n Complexity file level info : \n  %s" %json.dumps(res_dict,indent=4))
        return json.dumps(res_dict)
    except Exception as exception:
        logger.error("Exception in estimator()")
        logger.error(exception)
        exit()

if (__name__ == '__main__'):
    parser=argparse.ArgumentParser(description='Enter input Zip file path and output directory')
    parser.add_argument('--input', dest='input', type=str, help='Input Zip file path', required=True)
    parser.add_argument('--output', dest='output', type=str, help='Output directory path', required=True)
    parser.add_argument('--config', dest='config', type=str, help='Config file path', required=True)
    parser.add_argument('--bucket', dest='bucket', type=str, help='S3 bucket name', required=False)
    parser.add_argument('--topic', dest='topic', type=str, help='Topic ARN', required=False)

    args = parser.parse_args()
    if (not args.input or not args.output):
        logger.error('Please provide input and output directories')
        exit()

    print("\nUnzipping File and creating Folder Structure in Output directory...")
    output_dir=unzip_file(args.input.strip(), args.output.strip())
    #print(output_dir)
    print("Done! \nCalculating complexity for each file involved...")
    result=estimator(output_dir,args.config.strip()) 
    print("Done!\nWriting to Estimator Template file...") 
    SummaryPath, filePath=write_to_estimator_template(result,args.input.strip())
    #Uncomment below to upload file to S3
    if args.bucket:
        print("Done\nUploading file to S3")
        message=upload_to_s3(args.bucket.strip(), SummaryPath, filePath)
        if args.topic:
            print("Done\nSending Notification")
            send_notification(message, 'hadoop', args.topic.strip())
    print("Done!\nCheck out more info in log file: %s \n" %file_name)

            
