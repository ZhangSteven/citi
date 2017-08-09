# coding=utf-8
# 
# The read file utility functions needed by both open_citi.py and lookup.py
# 

from .utility import logger



def read_holding(ws, fields, row, column):
	"""
	Read holding from worksheet ws, starting at (row, column)
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



def read_position(ws, row, column, fields):
	"""
	Read a position on a particular row, starting at column
	"""
	logger.debug('read_position(): at row {0} column {1}'.format(row, column))
	position = {}
	for field in fields:
		position[field] = ws.cell_value(row, column)
		column = column + 1

	return position



def read_fields(ws, row, column):
	"""
	Read a list of fields from worksheet ws, starting at (row, column)
	"""
	fields = []
	while column < ws.ncols:
		cell_value = ws.cell_value(row, column)
		if ws.cell_value(row, column) == '':
			break

		fields.append(cell_value.strip())
		column = column + 1

	return fields