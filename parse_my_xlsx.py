#!/usr/bin/env python
import pandas as pd
import excel
import util

tables, names, headers, opts = excel.Excel.read('my_product.xlsx')
print(names)
# ['README', 'Product']
# We will read the 'Product' data sheet
t=tables[util.index('Product', names)]
t[:6].display()
#    Col1                              Col2                              Col3        Col4       Col5          Col6
#--  --------------------------------  --------------------------------  ----------  ---------  ------------  ------------
# 0  The real data starts from Row 3.
# 1
# 2  ProductID                         ProductName                       CategoryID  UnitPrice  UnitsInStock  Discontinued
# 3  1                                 Chai                              1           18         39            False
# 4  2                                 Chang                             1           19         17            False
# 5  3                                 Aniseed Syrup                     2           10         13            False

# The 3rd row is the header
header=t.loc[2]
# Real data starts from the 4th row
t=t[3:]
# fix the column header
t.columns=header
# reindex, so row index starts from 0, for convience
t.index=range(len(t))
# Now we have the right data
t[:3].display()
#      ProductID  ProductName      CategoryID    UnitPrice    UnitsInStock  Discontinued
#--  -----------  -------------  ------------  -----------  --------------  --------------
# 0            1  Chai                      1           18              39  False
# 1            2  Chang                     1           19              17  False
# 2            3  Aniseed Syrup             2           10              13  False

