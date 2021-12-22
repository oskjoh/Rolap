# -*- coding: utf-8 -*-
"""
Created on Wen Dec 22 09:16:03 2021

@author: Oskar Johansson
"""

import clr
import csv
import sys

clr.AddReferenceToFileAndPath (sys.argv[6])
clr.AddReference ("System.Data")

from Microsoft.AnalysisServices.AdomdClient import AdomdConnection , AdomdSchemaGuid
from System.Data import DataSet

def python_explore_schema(datasource, catalog, userid, password, fileout, field):
  connstring = "Data Source="+datasource+";Catalog="+catalog+";User ID="+userid+";Password="+password
  conn = AdomdConnection(connstring)
  conn.Open()
  ds = conn.GetSchemaDataSet(getattr(AdomdSchemaGuid, field), None)
  
  columnNames = [column.ColumnName for column in ds.Tables[0].Columns]
  columnClasses = [column.DataType.Name for column in ds.Tables[0].Columns]
  
  rows = []
  for row in ds.Tables[0].Rows:
    rows.append([str(x).encode('utf-8') for x in row])
  
  with open(fileout, 'w') as f:
    f.write('|'.join([(col.encode('utf-8')) for col in columnNames]))
    f.write('\n')
    output = '\n'.join(['|'.join(map(str,item)) for item in rows])
    f.write('\n')
    output = '\n'.join(['|'.join(map(str,item)) for item in rows])
    f.write(output)

python_explore_schema(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[7])
