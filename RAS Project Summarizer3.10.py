'''''
# !usr/bin/env python
# -*- coding: utf-8 -*-

# RAS Project Summarizer originally created
# 9/30/2019
# Adam Price
# pigankle@gmail.com
#
# PURPOSE: This script will analyze the project files for a RAS model, 
# and the plans, geometries and flow files in the same folder.  The
# output is an excel file with a table showing which geometry and flow are
# associated with each plan, as well as any orphaned (unused) plans, flows, or
# geometries.
# 	As of 9/30/2019, quasi-unsteady flows, sediment plans, and water
# quality plans are not supported
#
# HOW TO USE:
#   - Select your RAS project:
# 		- Double click on this script and navigate to a .prj file
# 		OR
# 		- Change the "fullpath" variable below.  If the .prj filename matches the
#       enclosing folder, you may provide the path of the enclosing folder.
#       If the .prj has a different name from the enclosing folder, you must 
#       provide the full path to the prj.
# 	- If your python installation has the xlsxwriter module, an excel file
# 		will be created in the RAS project directory with a table summarizing
# 		the plans, geometries and flows.  If xlsxwriter is not available, a csv
# 		file will be created instead.
#
#Examples for manual entry of path name when running from an IDE:
#If the project file is `"D:\STYX_FLOODING\STYX_2D\STYX_2D.prj"`
#then either of these will work:
#   `fullpath = "D:\STYX_FLOODING\STYX_2D\"`
#   `fullpath = "D:\STYX_FLOODING\STYX_2D\STYX_2D.prj"`
#
# For the project file below, the file name does not match the enclosing directory, so the full path is required.
#   `fullpath = "D:\STYX_FLOODING\STYX_2D\AsphodelLevees.prj"`
'''

# import packages
import glob
import pandas as pd
import os.path
import time
from datetime import datetime as dt
import numpy as np
# As of 2019, the default python installation for many USACE engineers is the one which comes with ArcGIS.
# Unfortunately, this sometimes does not have xslsxwriter available.
import imp

try:
    imp.find_module('xlsxwriter')
    import xlsxwriter

    xlwrfound = True
except ImportError:
    xlwrfound = False
import sys

from tkinter import filedialog, Tk

# Launch an open dialog to select the RAS project
root = Tk()
root.withdraw()
fullpath = filedialog.askopenfilename(parent=root, initialdir='c:',
                                        filetypes=(('RAS Project File', ".prj"),),
                                        title='Select a RAS .prj file')

# If the script is run from an IDE, the file path can be manually entered
#   fullpath = "D:\STYX_FLOODING\STYX_2D\AsphodelLevees.prj"

# Parse fullpath.
if fullpath[-4:] == ".prj":
    projName = os.path.basename(fullpath)
    os.chdir(os.path.dirname(fullpath))
else:  # This will be used only when the script is run from an IDE and the directory is manually entered.
    os.chdir(fullpath)
    projName = os.path.split(os.getcwd())[1] + ".prj"
dirpath = os.getcwd()

# create filename for storing output

if xlwrfound:
    sumFileName = projName[:-4] + " Project Summary_" + time.strftime("%Y-%m-%d_%H%M") + ".xlsx"
else:
    sumFileName = projName[:-4] + " Project Summary_" + time.strftime("%Y-%m-%d_%H%M") + ".csv"

# %%
print("Processing " + projName + " from the directory " + dirpath)
print("Output will be stored in " + sumFileName)

# initialize lists and dictionaries

flow_dict = {}
plan_dict = {}
geom_dict = {}
activeGeoms = []
activePlans = []
activeFlows = []
usedFlows = []
usedGeom = []

# PRJ
# Read all active flow, plan and geometry files from $$.prj
# Build lists of the active files
with open(projName, 'r') as f:
    for line in f:
        if line.startswith("Geom File"):
            activeGeoms.append(line[-4:].strip())
        if line.startswith("Plan File"):
            activePlans.append(line[-4:].strip())
        if line.startswith("Unsteady File"):
            activeFlows.append(line[-4:].strip())

# FLOW
# Loop through flow files (*.u## )
# Store the Flow Title and Description
flowTypes = ("*.u[0-9][0-9]", "*.f[0-9][0-9]")
# QuasiUnsteadyFlowFiles use a very different format and are not implemented, "*.q[0-9][0-9]")
flowFileList = []
for ft in flowTypes:
    flowFileList.extend(glob.glob(ft))

for filename in flowFileList:
    with open(filename, 'r') as f:
        timeStamp = dt.fromtimestamp(os.path.getmtime(filename)).strftime('%Y-%m-%d %H:%M')
        flowKey = filename[-3:]
        info = []
        for i in range(4):
            thisLine = next(f).strip()
            info.append(thisLine)
        flowDescr = "-"  # Sometimes the flow description is missing
        if info[2].startswith("BEGIN FILE DESCRIPTION"):
            flowDescr = info[3]
        flow_dict[flowKey] = flowKey, info[0][11:], flowDescr, timeStamp
    f.closed

# GEOMETRIES (*.g## )
for filename in glob.glob("*.g[0-9][0-9]"):
    with open(filename, 'r') as f:
        timeStamp = dt.fromtimestamp(os.path.getmtime(filename)).strftime('%Y-%m-%d %H:%M')
        geomKey = filename[-3:]
        info = []
        for i in range(6):
            thisLine = next(f).strip()
            info.append(thisLine)
        geoDescr = "-"
        if info[4].startswith("BEGIN GEOM DESCRIPTION"):
            geoDescr = info[5]
        geom_dict[geomKey] = geomKey, info[0][11:], geoDescr, timeStamp
    f.closed

# PLAN
# Loop through plan files.
for filename in glob.glob(projName[:-4] + ".p[0-9][0-9]"):
    with open(filename, 'r') as f:
        timeStamp = dt.fromtimestamp(os.path.getmtime(filename)).strftime('%Y-%m-%d %H:%M')
        planKey = filename[-3:]
        info = []
        planDescr = "-"
        storeDescr = False
        for i in range(31):
            thisLine = next(f).strip()
            info.append(thisLine)
            if info[i].startswith("END DESCRIPTION"):
                storeDescr = False
            if storeDescr:
                planDescr = planDescr + (info[i].strip())
            if info[i].startswith("BEGIN DESCRIPTION"):
                storeDescr = True

        # Get the flow and geometry files associated with this plan
        geomKey, flowKey = info[4][-3:], info[5][-3:]

        # Add flow and geometry files to the "used" lists if they are not already there.
        if not geomKey in usedGeom:
            usedGeom.append(geomKey)
        if not flowKey in usedFlows:
            usedFlows.append(flowKey)
        # Store the plan title, time window and Short ID
        entry = [info[0][11:], info[2][17:], planDescr, info[3][16:], timeStamp]
        # Also store the associated flow and geometry information
        entry.extend(flow_dict[flowKey])
        entry.extend(geom_dict[geomKey])
        # Note whether the Plan, Flow and Geometry are currently listed active in the $.prj file
        entry.extend([planKey in activePlans, geomKey in activeGeoms, flowKey in activeFlows])
        # Add to dictionary
        plan_dict[planKey] = entry
    f.closed

# Convert the plan dictionary to a dataframe
dfPlan = (pd.DataFrame.from_dict(plan_dict)).T
dfPlan.reset_index(inplace=True)  # .
dfPlan.columns = ["Plan #", "Plan Title", "Plan ShortID", "Plan Descr", "Time Range", "Plan Mod Time",
                  "Load #", "Load Name", "Load Descr", "Flow Mod Time",
                  "Geom #", "Geom Name", "Geom Descr", "Geom Mod Time",
                  "Plan in .prj?", "Geom in .prj?", "Flow in .prj?"]

#
# There may be flow or geometry files that are hanging out in the folder but are not being used by any plans.
# Add unused flow files to dataframe
#
for flowKey in flow_dict.keys():
    if flowKey not in usedFlows:
        dfPlan = dfPlan.append({"Load #": flowKey,
                                "Load Name": flow_dict[flowKey][1],
                                "Load Descr": flow_dict[flowKey][2],
                                "Flow Mod Time": flow_dict[flowKey][3],
                                "Flow in .prj?": False},
                               ignore_index=True)

# Add unused geom files to dataframe
for geomKey in geom_dict.keys():
    if geomKey not in usedGeom:
        dfPlan = dfPlan.append({"Geom #": geomKey,
                                "Geom Name": geom_dict[geomKey][1],
                                "Geom Descr": geom_dict[geomKey][2],
                                "Geom Mod Time": geom_dict[geomKey][3],
                                "Geom in .prj?": False},
                               ignore_index=True)

print("Processed " + str(len(plan_dict)) + " plan files")
print("Found " + str(len(flow_dict)) + " flow files.  " + str(
    len(flow_dict) - len(usedFlows)) + " are not used by any plans.")
print("Found " + str(len(geom_dict)) + " geometry files.  " + str(
    len(geom_dict) - len(usedGeom)) + " are not used by any plans.")

dfPlan = dfPlan.replace(np.nan, "-", regex=True)

if xlwrfound:
    workbook = xlsxwriter.Workbook(os.path.join(dirpath, sumFileName), options={'nan_inf_to_errors': False})
    worksheet = workbook.add_worksheet('sheet1')
    worksheet.add_table(0, 0, dfPlan.shape[0], dfPlan.shape[1] - 1,
                        {'data': dfPlan.values.tolist(),
                         'columns': [{'header': c} for c in dfPlan.columns.tolist()],
                         'style': 'Table Style Medium 9'})
    workbook.close()
else:
    export_csv = dfPlan.to_csv(os.path.join(dirpath, sumFileName), index=None, header=True)

print("Project summary written to " + os.path.join(dirpath, sumFileName))
print("Press <enter> to close this window.")
