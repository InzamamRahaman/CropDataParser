import wget
import os
import argparse
import wget
import urllib


days = range(1, 32)
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
years = range(2014, 2018)
#base_url = 'http://www.namistt.com/DocumentLibrary/Market%20Reports/Daily/Norris%20Deonarine%20NWM%20Daily%20Market%20Report%20-%20{0}%20{1}%20{2}.xls'

base_url = 'http://www.namistt.com/DocumentLibrary/Market Reports/Daily/Norris Deonarine NWM Daily Market Report - {0} {1} {2}.xls'


parser = argparse.ArgumentParser()
parser.add_argument('--output', help='The directory to download the files', default='input')
args = parser.parse_args()

for day in days:
    for month in months:
        for year in years:
            day_string = '{:02d}'.format(day)
            url = base_url.format(day_string, month, year)
            try:
                wget.download(url, args.output)
                print('Retrieving file for {0} {1} {2}'.format(day, month, year))
            except:
                print('Could not get ', url)
                #exit()







