# -*- coding: utf-8 -*-
"""
Created on Tue Jan 25 17:26:03 2022

@author: Oskar Johansson
"""

from read_olap import *

connstring = "Data Source="+sys.argv[1]+";Catalog="+sys.argv[2]+";User ID="+sys.argv[3]+";Password="+sys.argv[4]

if(sys.argv[8] == "read_olap"):
	dataset = connect_and_read_olap(connstring, sys.argv[7])


elif(sys.argv[8] == "explore_schema"):
	dataset = connect_and_explore_schema(connstring, sys.argv[7])

export_csv(dataset, sys.argv[5])
