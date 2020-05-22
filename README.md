# citi
Convert Citibank custodian files to Geneva format


## Limitations

1. Only CITI bank cash is written to the output cash file. UBS margin cash is on the cash section but with an empty currency string, currently ignored.



## ver 0.12@2017-10-18

1. Updated: citibank uses a new format, a new column "ISIN" is added, so the mapping from citibank id (actually SEDOL code) to ISIN is no longer needed, therefore we changed the open_citi.py code.

2. After this change, lookup.py and samples/InvestmentCodeLookup.xlsx are no longer needed, therefore deleted. Also the corresponding test code are updated.



## ver 0.1101@2017-8-16

1. Fixed bug: harded coded file path in test cases, failed in a different computer. Use os.path to workout the path.



## ver 0.11@2017-8-9

1. Updated samples/InvestmentCodeLookup.xlsx to make sure all code can be lookedup.
2. Updated logging.



## ver 0.1@2017-8-9

1. Generate cash and position csv output for star helios file on 2017-4-7.
2. Updated citi code to isin lookup table as of 2017-8-8.