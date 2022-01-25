# -*- coding: utf-8 -*-
"""
Created on Wen Dec 22 09:16:03 2021

@author: Oskar Johansson
"""
# Import required libraries

import clr
import csv
import sys

clr.AddReferenceToFileAndPath (sys.argv[6])
clr.AddReference ("System.Data")

from Microsoft.AnalysisServices.AdomdClient import AdomdConnection , AdomdDataAdapter , AdomdSchemaGuid
from System.Data import DataSet

# Define functions to connect and retrive data

def connect_and_read_olap(connstring, query):
	conn = AdomdConnection(connstring)
	conn.Open()
	cmd = conn.CreateCommand()
	cmd.CommandText = query
	adp = AdomdDataAdapter(cmd)
	dataset =  DataSet()
	adp.Fill(dataset)
	conn.Close()
	return dataset;

def connect_and_explore_schema(connstring, field):
	conn = AdomdConnection(connstring)
	conn.Open()
	dataset = conn.GetSchemaDataSet(getattr(AdomdSchemaGuid, field), None)
	conn.Close()
	return dataset;

# Function to export a csv of the retrieved olap data

def export_csv(dataset, fileout):
	# dataset hold your result as collection of tables
	# each tables has rows and each row has columns
	columnNames = [column.ColumnName for column in dataset.Tables[0].Columns]
	columnClasses = [column.DataType.Name for column in dataset.Tables[0].Columns]

	rows = []
	for row in dataset.Tables[0].Rows:
	    rows.append([str(x).encode('utf-8') for x in row])

	with open(fileout, 'w') as f:
		f.write('|'.join([(col.encode('utf-8')) for col in columnNames]))
		f.write('\n')
		output = '\n'.join(['|'.join(map(str,item)) for item in rows])
		f.write('\n')
		output = '\n'.join(['|'.join(map(str,item)) for item in rows])
		f.write(output)
