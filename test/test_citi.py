"""
Test the open_bal() method to open the trustee Macau Balanced Fund.
"""

import unittest2
import os, datetime
from xlrd import open_workbook
from citi.utility import get_current_directory
from citi.open_citi import open_citi, read_holding_fields, read_holding



class TestCiti(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCiti, self).__init__(*args, **kwargs)



    def test_read_holding_fields(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA 20170407.xls')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Holdings Report')
        fields = read_holding_fields(ws, 0, 1)
        self.assertEqual(len(fields), 17)
        self.assertEqual(fields[0], 'Asset Group')
        self.assertEqual(fields[1], 'Security ID')
        self.assertEqual(fields[2], 'Security Description')
        self.assertEqual(fields[3], 'Long/Short Indicator')
        self.assertEqual(fields[4], 'Shares/Par')
        self.assertEqual(fields[5], 'Curr')
        self.assertEqual(fields[15], 'Accounting Price  (Local CCY)')
        self.assertEqual(fields[16], 'FX Rate')



    def test_read_holding(self):
        file_name = os.path.join(get_current_directory(), 'samples', 'STA 20170407.xls')
        wb = open_workbook(filename=file_name)
        ws = wb.sheet_by_name('Holdings Report')
        port_values = {}
        read_holding(ws, port_values)
        holding = port_values['holding']
        self.assertEqual(len(holding), 22)
        self.verify_position1(holding[0])
        self.verify_position2(holding[1])
        self.verify_position3(holding[21])



    def verify_position1(self, position):
        """
        Verify the first postion in samples/STA 20170407.xls
        """
        self.assertEqual(len(position), 17)
        self.assertEqual(position['Asset Group'], 'BONDS')
        self.assertEqual(position['Security ID'], 'BDC4MV5')
        self.assertEqual(position['Security Description'], 'LENOVO PERPETUAL LENOVO 5 3/8 PERP')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 6800000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 6810800)
        self.assertAlmostEqual(position['Original Cost (Base)'], 46955110.64)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 6908324)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 101.593)
        self.assertAlmostEqual(position['FX Rate'], 6.9068882396)



    def verify_position2(self, position):
        """
        Verify the second postion in samples/STA 20170407.xls
        """
        self.assertEqual(len(position), 17)
        self.assertEqual(position['Asset Group'], '')
        self.assertEqual(position['Security ID'], 'BDF16K0')
        self.assertEqual(position['Security Description'], 'HUARONG FIN II HRAM 4 5/8 06/03/26')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 2000000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 2030000)
        self.assertAlmostEqual(position['Original Cost (Base)'], 13913544.12)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 2021662)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 101.0831)
        self.assertAlmostEqual(position['FX Rate'], 6.9068882396)



    def verify_position3(self, position):
        """
        Verify the last postion in samples/STA 20170407.xls
        """
        self.assertEqual(len(position), 17)
        self.assertEqual(position['Asset Group'], '')
        self.assertEqual(position['Security ID'], 'XS1587894343')
        self.assertEqual(position['Security Description'], 'TEWOO GROUP TEWOOG 4 5/8 04/06/20')
        self.assertEqual(position['Long/Short Indicator'], 'L')
        self.assertAlmostEqual(position['Shares/Par'], 2400000)
        self.assertEqual(position['Curr'], 'USD')
        self.assertAlmostEqual(position['Original Cost (Local)'], 2383464)
        self.assertAlmostEqual(position['Original Cost (Base)'], 16380070.1)
        self.assertEqual(position['Amortized Cost (Local)'], '')
        self.assertAlmostEqual(position['Position Accounting Market Value (Local CCY)'], 2419219.2)
        self.assertAlmostEqual(position['Accounting Price  (Local CCY)'], 100.8008)
        self.assertAlmostEqual(position['FX Rate'], 6.9068882396)