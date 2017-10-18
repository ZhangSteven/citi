"""
Test the open_bal() method to open the trustee Macau Balanced Fund.
"""

import unittest2
import os, datetime
from xlrd import open_workbook
from citi.utility import get_current_directory
from citi.open_citi import open_citi, read_grand_total, update_cash_data, \
                            get_portfolio_date
from citi.read_file import read_fields, read_holding



class TestCiti(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCiti, self).__init__(*args, **kwargs)



    def test_read_fields(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA1_20171017.xlsx')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Holdings Report ISIN')
        fields = read_fields(ws, 0, 1)
        self.assertEqual(len(fields), 23)
        self.assertEqual(fields[0], 'Asset Group')
        self.assertEqual(fields[1], 'Security ID')
        self.assertEqual(fields[2], 'ISIN')
        self.assertEqual(fields[3], 'Security Description')
        self.assertEqual(fields[4], 'Long/Short Indicator')
        self.assertEqual(fields[5], 'Shares/Par')
        self.assertEqual(fields[6], 'Curr')
        self.assertEqual(fields[16], 'Accounting Price  (Local CCY)')
        self.assertEqual(fields[17], 'FX Rate')



    def test_read_grand_total(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA1_20171017.xlsx')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Holdings Report ISIN')
        fields = read_fields(ws, 0, 1)
        self.assertAlmostEqual(read_grand_total(ws, 0, 1, fields, 'Shares/Par'), 137080000)



    def test_read_holding(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA1_20171017.xlsx')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Holdings Report ISIN')
        fields = read_fields(ws, 0, 1)
        holding = read_holding(ws, fields, 1, 1)
        self.assertEqual(len(holding), 51)
        self.verify_position1(holding[0])
        self.verify_position2(holding[1])
        self.verify_position3(holding[50])



    def test_read_cash(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA1_20171017.xlsx')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Accrued Interest on Cash Accoun')
        fields = read_fields(ws, 0, 1)
        cash = update_cash_data(read_holding(ws, fields, 1, 1))
        self.assertEqual(len(cash), 1)
        self.verify_cash(cash[0])



    def test_open_citi(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA1_20171017.xlsx')
        port_values = {}
        output_dir = os.path.join(get_current_directory(), 'samples')
        file_list = open_citi(file_name, port_values, output_dir, 'star_helios_')
        holding = port_values['holding']
        self.assertEqual(len(holding), 51)
        self.verify_position1(holding[0])
        self.verify_position2(holding[1])
        self.verify_position3(holding[50])

        cash = port_values['cash']
        self.assertEqual(len(cash), 1)
        self.verify_cash(cash[0])

        self.assertEqual(get_portfolio_date(port_values), datetime.datetime(2017,10,17))
        self.assertEqual(port_values['portfolio_id'], '40001')
        
        self.assertEqual(len(file_list), 2)
        self.assertEqual(file_list[0], os.path.join(get_current_directory(), 'samples', 'star_helios_2017-10-17_cash.csv'))
        self.assertEqual(file_list[1], os.path.join(get_current_directory(), 'samples', 'star_helios_2017-10-17_position.csv'))



    def verify_position1(self, position):
        """
        Verify the first postion in samples/STA1_20171017.xlsx
        """
        self.assertEqual(len(position), 23)

        self.assertEqual(position['Asset Group'], 'BONDS')
        self.assertEqual(position['Security ID'], 'BCRW3Z8')
        self.assertEqual(position['ISIN'], 'USG24524AH67')
        self.assertEqual(position['Security Description'], 'COUNTRY GARDEN COGARD 7 1/4 04/04/21')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 4000000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 4154000)
        self.assertAlmostEqual(position['Original Cost (Base)'], 28254114.72)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 4170012)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 104.2503)
        self.assertAlmostEqual(position['FX Rate'], 6.6161632869)



    def verify_position2(self, position):
        """
        Verify the second postion in samples/STA1_20171017.xlsx
        """
        self.assertEqual(len(position), 23)
        self.assertEqual(position['Asset Group'], '')
        self.assertEqual(position['Security ID'], 'BDDWMY1')
        self.assertEqual(position['ISIN'], 'USG2120QAC09')
        self.assertEqual(position['Security Description'], 'CHINASOUTH POWER SOPOWZ 3 1/2 05/08/27')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 2000000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 2051460)
        self.assertAlmostEqual(position['Original Cost (Base)'], 13537415.86)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 2044050)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 102.2025)
        self.assertAlmostEqual(position['FX Rate'], 6.6161632869)



    def verify_position3(self, position, security_id_updated=False):
        """
        Verify the last postion in samples/STA1_20171017.xlsx
        """
        self.assertEqual(len(position), 23)
        self.assertEqual(position['Asset Group'], '')
        self.assertEqual(position['Security ID'], 'XS1688369617')
        self.assertEqual(position['ISIN'], 'XS1688369617')
        self.assertEqual(position['Security Description'], 'KAISA GROUP KAISAG 8 1/2 06/30/22')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 2000000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 2006000)
        self.assertAlmostEqual(position['Original Cost (Base)'], 13201711.09)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 2008140)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 100.407)
        self.assertAlmostEqual(position['FX Rate'], 6.6161632869)



    def verify_cash(self, position):
        """
        Verify the cash position in samples/STA1_20171017.xlsx
        """
        self.assertEqual(len(position), 8)
        self.assertEqual(position['Local CCY'], 'USD')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 9404448.39)
        self.assertAlmostEqual(position['Accrued Interest'], 0)
        self.assertAlmostEqual(position['Exchange Rate'], 0.151145)
        self.assertAlmostEqual(position['Accounting Market Value (VCY)'], 62221366.17)
        self.assertEqual(position['As Of'], datetime.datetime(2017,10,17))