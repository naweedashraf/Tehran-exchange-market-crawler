import os, csv, openpyxl, logging, jdatetime, argparse

def getFileNames(folder):
    try:
        # folder path
        dir_path = r'./'+folder+''

        # list to store files
        res = []

        # Iterate directory
        for path in os.listdir(dir_path):
            # check if current path is a file
            if os.path.isfile(os.path.join(dir_path, path)):
                res.append(path)

        logging.info('file names acquired '+str(jdatetime.datetime.now()))
        return res
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tgetFileNames function\n')
        return []

def deleteFiles(sourceFolder, files):
    try:
        for file in files:
            os.remove('./'+sourceFolder+'/'+file)
            logging.info(file + ' deleted '+str(jdatetime.datetime.now()))
        
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tdeleteFiles function\n')

def changeExcelToCsv(sourceFolder, destinationFolder, isRemoving):
    try:
        #create destination folder if it not exists
        if destinationFolder not in os.listdir('./'):
            os.mkdir('./'+destinationFolder) 
        
        #get xlsx file names
        files = getFileNames(sourceFolder)
        emptyFiles = []
        workingDayFiles = []

        #move xlsx file contents to csv file
        for item in files:
            records = []
            workbook = openpyxl.load_workbook('./' + sourceFolder + '/'+ item)
            sheet = workbook.active

            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                record = []
                for cell in row:
                    record.append(cell.value)
                records.append(record)

            if len(records)>3:
                workingDayFiles.append(item)
                with open('./' + destinationFolder + '/'+ item[:-4] + 'csv', 'w', newline='',encoding='utf-8') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerows(records)
            else:
                emptyFiles.append(item)

        deleteFiles(sourceFolder,emptyFiles)
        
        if isRemoving:
            deleteFiles(sourceFolder,workingDayFiles)
        
        logging.info('excel files data transfered to csv files '+str(jdatetime.datetime.now()))

    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tchangeExcelToCsv function\n')



if __name__ == "__main__":
    logging.basicConfig(filename='info.log', level=logging.DEBUG)
    
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument('source', type=str, help='source folder of excel files')
        parser.add_argument('is_removing', type=str, help='set this to True if you want to remove excel files after transfering data to csv files, False value will keep the files')
        args = parser.parse_args()
        
        if args.is_removing == 'True':
            changeExcelToCsv(args.source,'datalake',True)
        elif args.is_removing == 'False':
            changeExcelToCsv(args.source,'datalake',False)
        else:
            raise Exception('is_removing value is not correct.')
        
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tmain function|\n')    
