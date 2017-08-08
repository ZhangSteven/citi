# coding=utf-8
# 
# Open the CitiBank custodian file, read position and cash from it, save to
# output csv files for Geneva reconciliation.
# 

from .utility import logger


class InconsistentGrandTotal(Exception):
	pass



def open_citi(filename, port_values, output_dir, output_prefix):
	"""
	Read a citibank excel file, convert them to geneva format (csv),
	then return the csv file names (full path).
	"""
	logger.info('open_citi(): {0}'.format(filename))
	
	wb = open_workbook(filename=file_name)
	ws = wb.sheet_by_name('Holdings Report')
	read_holding(ws, port_values)
	validate_holding(ws, port_values)

	# cash reading

	# write output csv



def read_holding(ws, port_values):
	"""
	Read holding 
	"""
	column = 1	# fields start at column 1 (first column 0)
	fields = read_holding_fields(ws, 0, column)
	
	row = 1		# positions start at row 1 (first row is 0)
	holding = []
	port_values['holding'] = holding
	while row < ws.nrows:
		if ws.cell_value(row, 2) == '':	# the security id field
			break

		holding.append(read_position(ws, row, column, fields))
		row = row + 1
	# end of while loop

	validate_holding(port_values, read_grand_total(ws, 0, column, fields))



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



def read_holding_fields(ws, row, column):
	fields = []
	while column < ws.ncols:
		cell_value = ws.cell_value(row, column)
		if ws.cell_value(row, column) == '':
			break

		fields.append(cell_value.strip())
		column = column + 1

	return fields



def validate_holding(port_values, total_shares_par):
	"""
	Read the grand total numbers in the holding section, use this to
	validate the holding.
	"""
	total_quantity = 0
	for position in port_values['holding']:
		total_quantity = total_quantity + position['Shares/Par']

	if abs(total_quantity - total_shares_par) > 0.01:
		logger.error('validate_holding(): calculated total quantity {0} \
			is different from grand total {1}'.format(total_quantity, total_shares_par))
		raise InconsistentGrandTotal()



def read_grand_total(ws, row, column, fields):
	"""
	Read the grand total number of shares/par for the holdings.

	row: start to search grand total from which row
	column: which column does the fields start
	"""
	while row < ws.nrows:
		cell_value = ws.cell_value(row, 0)
		if isinstance(cell_value, str) and cell_value.startswith('Grand Total'):
			for field in fields:
				if field == 'Shares/Par':
					return ws.cell_value(row, column)

				column = column + 1

		row = row + 1
