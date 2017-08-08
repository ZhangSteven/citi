# coding=utf-8
# 
# Open the CitiBank custodian file, read position and cash from it, save to
# output csv files for Geneva reconciliation.
# 

from .utility import logger, get_datemode
from xlrd import open_workbook, xldate



class InconsistentGrandTotal(Exception):
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
	port_values['cash'] = map_cash_date(read_holding(ws, fields, 1, 1))
	validate_holding(port_values['cash'], ws, 0, 1, fields, 'Accounting Market Value (VCY)')

	return write_csv(port_values, output_dir, output_prefix)



def get_portfolio_id(ws):
	return ''



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



def map_cash_date(cash_accounts):
	for account in cash_accounts:
		account['As Of'] = xldate.xldate_as_datetime(account['As Of'], get_datemode())

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
	pass



def write_holding_csv(port_values, output_dir, output_prefix):
	pass