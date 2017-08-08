# coding=utf-8
# 
# Open the CitiBank custodian file, read position and cash from it, save to
# output csv files for Geneva reconciliation.
# 

from .utility import logger



def open_citi(filename, port_values, output_dir, output_prefix):
	"""
	Read a citibank excel file, convert them to geneva format (csv),
	then return the csv file names (full path).
	"""
	logger.info('open_citi(): {0}'.format(filename))
	
	wb = open_workbook(filename=file_name)
	ws = wb.sheet_by_name('Holdings Report')
	read_holding(ws, port_values)
	# validate_holding(port_values)

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



def read_position(ws, row, column, fields):
	"""
	Read a position on a particular row
	"""
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


