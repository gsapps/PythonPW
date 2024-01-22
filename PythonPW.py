import warnings

import os
import glob
from datetime import datetime

import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'

import PySimpleGUI as psg
import xlsxwriter

dataFolder = 'C:\\Users\\billo\\Downloads'
filenamePattern = "orders-paypwrvcw-weebly*.csv"
neededColumns = [
	"Order #",
	"Date",
	"Status",
	"Currency",
	"Subtotal",
	"Shipping First Name",
	"Shipping Last Name",
	"Shipping Email",
	"Shipping Phone",
	"Billing Postal Code",
	"Billing Country",
	"Product Name",
	"Order Notes"
	]

def PW():
	# get the latest "orders-paypwrvcw-weebly*.csv" file
	filenames = glob.glob(os.path.join(dataFolder, filenamePattern))
	filenames.sort(key=os.path.getmtime, reverse=True)
	filename = filenames[0]  # look only at the latest
	df = pd.read_csv(filename, header=0)[neededColumns]

	# each weebly order comes as two separate lines in the .csv (with the same Order Number).
	# The first line contains almost all of the needed info.
	# The second line contains only the Product Name.
	# Here, we shift the column values up to be in the first line of the two
	shiftUpColumn = df['Product Name'].tolist()
	shiftUpColumn.pop(0)
	shiftUpColumn.append('dummy')
	df['Product Name'] = shiftUpColumn
	# then delete the second lines, which we now don't need
	df.drop(df[df.Status != 'paid'].index, inplace=True)

	# Export the new file with a name of the form: Orders 11-2023.xlsx
	date = datetime.strptime(df['Date'][0], '%Y/%m/%d')  # all orders are in the same month/year, so get any row
	exportFilename = 'Orders {0:02n}-{1}.xlsx'.format(date.month, date.year)
	exportFilename = os.path.join(dataFolder, exportFilename)

	# change the dates to desired format
	df['Date'] = df['Date'].map(lambda str: datetime.strptime(str, '%Y/%m/%d').strftime('%m/%d/%Y'))

	try:
		with pd.ExcelWriter(exportFilename, engine='xlsxwriter') as writer:
			df.to_excel(writer, index=False)
			worksheet = writer.sheets['Sheet1'].autofit() # currently, xlsxwrite is the only way to get autofit
	except:
		psg.Popup('The file {0} is currently in use.'.format(exportFilename))
		return

	psg.popup_auto_close('File created: ' + exportFilename)

	
PW()