# coding=utf-8
# 
# Open the CitiBank custodian file, read position and cash from it, save to
# output csv files for Geneva reconciliation.
# 

from .utility import logger, get_datemode, convert_datetime_to_string, \
						get_csv_file_name
from xlrd import open_workbook, xldate
import csv, os


class InconsistentGrandTotal(Exception):
	pass

class InvalidPortfolioName(Exception):
	pass



def open_citi(filename, port_values, output_dir, output_prefix):
	"""
	Read a citibank excel file, convert them to geneva format (csv),
	then return the csv file names (full path).
	"""
	logger.info('open_citi(): {0}'.format(filename))
	
	wb = open_workbook(filename=filename)
	ws = wb.sheet_by_name('Index Page')
	port_values['portfolio_id'] = get_portfolio_id(ws)

	ws = wb.sheet_by_name('Holdings Report')
	fields = read_fields(ws, 0, 1)
	port_values['holding'] = read_holding(ws, fields, 1, 1)
	validate_holding(port_values['holding'], ws, 0, 1, fields, 'Shares/Par')

	ws = wb.sheet_by_name('Accrued Interest on Cash Accoun')
	fields = read_fields(ws, 0, 1)
	port_values['cash'] = update_cash_data(read_holding(ws, fields, 1, 1))
	validate_holding(port_values['cash'], ws, 0, 1, fields, 'Accounting Market Value (VCY)')

	return write_csv(port_values, output_dir, output_prefix)



def get_portfolio_id(ws):
	"""
	Get the portfolio name from sheet "Index Page" and map it to a 
	portfolio id.
	"""
	logger.debug('get_portfolio_id()')
	row = 0
	while row < ws.nrows:
		if ws.cell_value(row, 2) == 'Account:':
			break

		row = row + 1
	# end of while loop

	if ws.cell_value(row, 3).strip() == 'STA1 - STAR HELIOS PLC-CHINA LIFE':
		return '40001'
	else:
		logger.error('get_portfolio_id(): invalid portfolio name: {0}'.format(ws.cell_value(row, 3).strip()))
		raise InvalidPortfolioName()



def read_holding(ws, fields, row, column):
	"""
	Read holding 
	"""
	holding = []
	while row < ws.nrows:
		if ws.cell_value(row, 2) == '':	# the first field can be empty
										# for bond positions
			break

		holding.append(read_position(ws, row, column, fields))
		row = row + 1
	# end of while loop

	return holding



def update_cash_data(cash_accounts):
	"""
	Update certain cash data to other format.

	Local CCY: change to standard representation such as USD, HKD, etc.
	As Of: change to python datetime format.
	"""
	logger.debug('update_cash_data(): start')
	c_map = {
		'US DOLLAR':'USD'
	}

	for account in cash_accounts:
		logger.debug('update_cash_data(): {0}, amount {1}'.\
						format(account['Local CCY'], account['Position Accounting Market Value (Local CCY)']))
		account['As Of'] = xldate.xldate_as_datetime(account['As Of'], get_datemode())
		try:
			account['Local CCY'] = c_map[account['Local CCY']]
		except KeyError:
			logger.error('update_cash_data(): failed to map {0} to standard representation'.format(account['Local CCY']))
			raise

	return cash_accounts



def read_position(ws, row, column, fields):
	"""
	Read a position on a particular row
	"""
	logger.debug('read_position(): at row {0} column {1}'.format(row, column))
	position = {}
	for field in fields:
		position[field] = ws.cell_value(row, column)
		column = column + 1

	return position



def read_fields(ws, row, column):
	fields = []
	while column < ws.ncols:
		cell_value = ws.cell_value(row, column)
		if ws.cell_value(row, column) == '':
			break

		fields.append(cell_value.strip())
		column = column + 1

	return fields



def validate_holding(holding, ws, row, column, fields, key_field):
	"""
	Read the grand total number for the key_field in the holding section, 
	then use that number to validate the holding.
	"""
	total = 0
	for position in holding:
		total = total + position[key_field]

	grand_total = read_grand_total(ws, row, column, fields, key_field)
	if abs(total - grand_total) > 0.01:
		logger.error('validate_holding(): calculated total {0} \
			is different from grand total {1}'.format(total, grand_total))
		raise InconsistentGrandTotal()



def read_grand_total(ws, row, column, fields, key_field):
	"""
	Read the grand total number of a key field based on the holdings.

	row: start to search grand total from which row
	column: which column does the fields start
	"""
	while row < ws.nrows:
		cell_value = ws.cell_value(row, 0)
		if isinstance(cell_value, str) and cell_value.startswith('Grand Total'):
			for field in fields:
				if field == key_field:
					return ws.cell_value(row, column)

				column = column + 1

		row = row + 1



def get_portfolio_date(port_values):
	"""
	The date of holdings and cash data. Here we assume the date of the cash
	entries are the same and represent the date of the holdings.
	"""
	return port_values['cash'][0]['As Of']



def create_csv_file_name(date_string, output_dir, file_prefix, file_suffix):
	"""
	Create the output csv file name based on the date string, as well as
	the file suffix: cash, afs_positions, or htm_positions
	"""
	csv_filename = "".join([file_prefix, date_string, '_', file_suffix, '.csv'])
	return os.path.join(output_dir, csv_filename)



def write_csv(port_values, output_dir, output_prefix):
	cash_file = write_cash_csv(port_values, output_dir, output_prefix)
	position_file = write_holding_csv(port_values, output_dir, output_prefix)
	return [cash_file, position_file]



def write_cash_csv(port_values, output_dir, output_prefix):
	portfolio_date = convert_datetime_to_string(get_portfolio_date(port_values))
	file_name = get_csv_file_name(output_dir, output_prefix+portfolio_date, 'cash')
	logger.debug('write_cash_csv(): {0}'.format(file_name))
	with open(file_name, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')
		fields = ['currency', 'balance']
		file_writer.writerow(['portfolio', 'custodian', 'date'] + fields)

		for position in port_values['cash']:
			row = [port_values['portfolio_id'], 'CITI', portfolio_date]
			for fld in fields:
				if fld == 'currency':
					item = position['Local CCY']
				elif fld == 'balance':
					item = position['Position Accounting Market Value (Local CCY)']

				row.append(item)

			file_writer.writerow(row)

	return file_name



def write_holding_csv(port_values, output_dir, output_prefix):
	portfolio_date = convert_datetime_to_string(get_portfolio_date(port_values))
	file_name = get_csv_file_name(output_dir, output_prefix+portfolio_date, 'position')
	logger.debug('write_holding_csv(): {0}'.format(file_name))
	with open(file_name, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')

		# except for name, all fields are mandatory to do a position recon
		# in Geneva
		fields = ['geneva_investment_id', 'isin', 'bloomberg_figi', 'name', 
					'currency', 'quantity']
		file_writer.writerow(['portfolio', 'custodian', 'date'] + fields)

		for position in port_values['holding']:
			row = [port_values['portfolio_id'], 'CITI', portfolio_date]
			for fld in fields:
				if fld == 'currency':
					item = position['Curr']
				elif fld == 'name':
					item = position['Security Description']
				elif fld == 'quantity':
					item = position['Shares/Par']
				else:
					try:
						item = position[fld]
					except KeyError:
						item = ''

				row.append(item)

			file_writer.writerow(row)
		# end of for loop

	return file_name