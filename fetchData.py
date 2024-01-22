import requests, argparse, jdatetime, os, logging
from datetime import timedelta

def getDates(date1, date2):
    try:
        #change date to gregorian
        startDate = jdatetime.date(int(date1.split('-')[0]),int(date1.split('-')[1]),int(date1.split('-')[2]) ).togregorian()
        endDate = jdatetime.date(int(date2.split('-')[0]),int(date2.split('-')[1]),int(date2.split('-')[2]) ).togregorian()
        if startDate > endDate:
            raise Exception("start date is bigger than end date")
        
        #store dates between start date and end date
        delta = timedelta(days=1)
        dates = []
        while startDate <= endDate:
            dates.append(startDate.isoformat())
            startDate += delta
        persianDates = []

        #change date to jalali date
        for date in dates:
            persianDates.append(jdatetime.date.fromgregorian(day=int(date[8:]),month=int(date[5:7]),year=int(date[:4])))

        logging.info('dates period calculated '+str(jdatetime.datetime.now()))
        return persianDates
    
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tgetDates function\n')
        return []


def getExcelFiles(date1,date2):
    try:
        dates = getDates(date1,date2)

        if 'stage' not in os.listdir('./'):
            os.mkdir('./stage')

        URL = "http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d="
        
        for date in dates:
            #exclude thursdays and fridays
            if date.weekday() < 5:
                response = requests.get(URL+str(date))
                open("./stage/MarketWatchPlus-"+str(date)+".xlsx", "wb").write(response.content)
                logging.info('MarketWatchPlus-'+str(date)+".xlsx file succesfully downloaded "+str(jdatetime.datetime.now()))

        logging.info('all files acquired '+str(jdatetime.datetime.now()))

    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tgetExcelFiles function|\n')


if __name__ == "__main__":
    logging.basicConfig(filename='info.log', level=logging.DEBUG)
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument('start_date', type=str, help='enter in the yyyy-mm-dd format')
        parser.add_argument('end_date', type=str, help='enter in the yyyy-mm-dd format')
        args = parser.parse_args()

        getExcelFiles(args.start_date,args.end_date)
    except Exception as e:
        with open("error.log", "a") as f:
            f.write(str(e)+'\t|\tmain function|\n')

