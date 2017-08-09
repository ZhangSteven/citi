# coding=utf-8
# 
# Lookup isin code based on citi security id.
# 

from .utility import logger, get_current_directory
from .read_file import read_holding, read_fields
from xlrd import open_workbook
import os, re



def initialize_isin_table():
	"""
	Read the InvestmentCodeLookup.xlsx file, create a lookup table for

	citi code -> isin code
	"""
	wb = open_workbook(filename=os.path.join(get_current_directory(), 
						'samples', 'InvestmentCodeLookup.xlsx'))
	ws = wb.sheet_by_name('Sheet1')
	fields = read_fields(ws, 0, 0)
	lines = read_holding(ws, fields, 1, 0)
	isin_map = {}
	for item in lines:
		isin_map[item['CITI code']] = item['ISIN/Geneva']

	return isin_map



# initialized the lookup table if it's not there
if not 'isin_map' in globals():
	isin_map = initialize_isin_table()



def lookup_isin_from_id(security_id):
	logger.debug('lookup_isin_from_id(): start')
	if is_isin(security_id):
		return security_id
	else:
		try:
			global isin_map
			return isin_map[security_id]
		except KeyError:
			logger.error('lookup_isin_from_id(): failed to lookup isin code for {0}'.format(security_id))
			return ''



def is_isin(security_id):
	if re.search('^[A-Z]{2}[A-Z0-9]{10}$', security_id) is None:
		return False
	else:
		return True