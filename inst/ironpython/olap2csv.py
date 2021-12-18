# -*- coding: utf-8 -*-
"""
Created on Sat Sep 21 14:14:26 2019

@author: ojo
"""
import clr
import csv
import sys

clr.AddReferenceToFileAndPath (sys.argv[6])
clr.AddReference ("System.Data")

from Microsoft.AnalysisServices.AdomdClient import AdomdConnection , AdomdDataAdapter
from System.Data import DataSet

def python_olap2csv(datasource, catalog, userid, password, fileout, query):
	connstring = "Data Source="+datasource+";Catalog="+catalog+";User ID="+userid+";Password="+password
	conn = AdomdConnection(connstring)
	conn.Open()
	cmd = conn.CreateCommand()
	cmd.CommandText = query
	adp = AdomdDataAdapter(cmd)
	datasetParam =  DataSet()
	adp.Fill(datasetParam)
	conn.Close();

	# datasetParam hold your result as collection a\of tables
	# each tables has rows
	# and each row has columns
	#for i in range(0,1):
	#	print datasetParam.Tables[0].Columns[i].ColumnName+","+
	columnNames = [column.ColumnName for column in datasetParam.Tables[0].Columns]
	columnClasses = [column.DataType.Name for column in datasetParam.Tables[0].Columns]

	rows = []
	for row in datasetParam.Tables[0].Rows:
	    rows.append([str(x).encode('utf-8') for x in row])
	#columnNames_fix = [item.replace("\\[u\\'", "\\'") for item in columnNames]
	#columnNames_fix = [item.replace("\\'\\]", "'") for item in columnNames_fix]

	with open(fileout, 'w') as f:
		#f.write('|'.join([col for col in columnClasses]))
		#f.write('\n')
		f.write('|'.join([(col.encode('utf-8')) for col in columnNames]))
		f.write('\n')
		output = '\n'.join(['|'.join(map(str,item)) for item in rows])
		f.write('\n')
		output = '\n'.join(['|'.join(map(str,item)) for item in rows])
		f.write(output)

python_olap2csv(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[7])		