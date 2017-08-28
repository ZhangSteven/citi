"""
Test the open_bal() method to open the trustee Macau Balanced Fund.
"""

import unittest2
from citi.lookup import isin_map, lookup_isin_from_id


class TestLookup(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestLookup, self).__init__(*args, **kwargs)



    def test_isin_map(self):
        self.assertTrue(len(isin_map) > 45)	# we will keep updating the lookup file
        									# so it's difficult to tell how many records
        									# there will be.
        self.assertEqual(isin_map['BF04Y37'], 'XS1562292026')
        self.assertEqual(isin_map['BF08G22'], 'XS1573134878')
        self.assertEqual(isin_map['BF282K0'], 'XS1572322409')
        self.assertEqual(isin_map['BYQCMB2'], 'XS1565684062')



    def test_lookup_isin_from_id(self):
    	# it's isin code, stay the same
    	self.assertEqual(lookup_isin_from_id('XS1572322409'), 'XS1572322409')

    	# it's not isin code, lookup
    	self.assertEqual(lookup_isin_from_id('BDC4MV5'), 'XS1575529539')

    	# it's not isin code, lookup failed
    	self.assertEqual(lookup_isin_from_id('<strange_code>'), '') 