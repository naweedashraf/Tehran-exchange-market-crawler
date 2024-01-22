import os, csv, logging, jdatetime

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

class TSETMC:
    def __init__(self, sourceFolderName):
        try:
            self.source = sourceFolderName
            self.records = []
            fileNames = []
            logging.info('Object created '+str(jdatetime.datetime.now()))
            
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\tinit function|\n')    
    
    def getFileNames(self):
        try:
            # folder path
            dir_path = r'./'+self.source+''

            # list to store files
            res = []

            # Iterate directory
            for path in os.listdir(dir_path):
                # check if current path is a file
                if os.path.isfile(os.path.join(dir_path, path)):
                    res.append(path)

            logging.info('file names acquired '+str(jdatetime.datetime.now()))
            self.fileNames = res[:]
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\tgetFileNames function\n')
            self.fileNames = []
    
    def addRecords(self):
        try:
            self.getFileNames()
            for item in self.fileNames:
                 with open('./' + self.source + '/'+ item, 'r', newline='',encoding='utf-8') as csvfile:
                    dummy = []
                    for line in csvfile:
                        dummy.append(line)
                    self.records.append(dummy)
        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\taddRecords function|\n')  
    
    def maxTradeVolume(self):
        try:
            result = []
            result.sort()
            for day in self.records:
                for i in range(len(day)):
                    if i >= 3:
                        if len(result)<10:
                            result.append([int(day[i].split(',')[3]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]])
                            result.sort()
                        elif int(day[i].split(',')[3]) > result[0][0]:
                            result[0] = [int(day[i].split(',')[3]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]]
                            result.sort()
            result.sort(reverse=True)
            logging.info('======================================================= ')
            logging.info('top 10 shares by total volum traded: ')
            logging.info('======================================================= ')
            for i in range(len(result)):
                logging.info(str(i) + '- stock symbol: '+result[i][3]+' - stock name: ' +result[i][1]+' - on ' + result[i][2] + ' :\t'+str(result[i][0])) 

        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\tmaxTotalVolume function|\n')   
        
    def maxPriceIncrease(self):
        try:
            result = []
            result.sort()
            for day in self.records:
                for i in range(len(day)):
                    if i >= 3:
                        if float(day[i].split(',')[12]) < 20:
                            if len(result)<10:
                                result.append([float(day[i].split(',')[12]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]])
                                result.sort()
                            elif float(day[i].split(',')[12]) > result[0][0]:
                                result[0] = [float(day[i].split(',')[12]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]]
                                result.sort()
            result.sort(reverse=True)
            logging.info('======================================================= ')
            logging.info('top 10 shares by maximum increase in close price(below 20 percent): ')
            logging.info('======================================================= ')
            for i in range(len(result)):
                logging.info(str(i) + '- stock symbol: '+result[i][3]+' - stock name: ' +result[i][1]+' - on ' + result[i][2] + ' :\t'+str(result[i][0])) 

        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\tmaxTotalVolume function|\n')   

    def maxPriceDecrease(self):
        try:
            result = []
            result.sort()
            for day in self.records:
                for i in range(len(day)):
                    if i >= 3:
                        if float(day[i].split(',')[12]) > -20:
                            if len(result)<10:
                                result.append([float(day[i].split(',')[12]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]])
                                result.sort(reverse=True)
                            elif float(day[i].split(',')[12]) < result[0][0]:
                                result[0] = [float(day[i].split(',')[12]),day[i].split(',')[1],day[1].split(' ')[4][:10],day[i].split(' ')[0]]
                                result.sort(reverse=True)
            result.sort()
            logging.info('======================================================= ')
            logging.info('top 10 shares by maximum decrease in close price(below 20 percent): ')
            logging.info('======================================================= ')
            for i in range(len(result)):
                logging.info(str(i) + '- stock symbol: '+result[i][3]+' - stock name: ' +result[i][1]+' - on ' + result[i][2] + ' :\t'+str(result[i][0])) 

        except Exception as e:
            with open("error.log", "a") as f:
                f.write(str(e)+'\t|\tmaxTotalVolume function|\n')   

if __name__ == "__main__":
    logging.basicConfig(filename='info.log', level=logging.DEBUG,encoding='utf-8')
    
    try:
        model = TSETMC('datalake')
        model.addRecords()
        model.maxTradeVolume()
        model.maxPriceIncrease()
        model.maxPriceDecrease()
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tmain function|\n')    
