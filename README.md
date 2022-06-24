# HEC-RAS-Project-Summarizer

_RAS Project Summarizer originally created on 9/30/2019 by  Adam Price pigankle@gmail.com_

PURPOSE: This script will analyze the project files for a RAS model, and the plans, geometries and flow files in the same folder.  The output is an excel file with a table showing which geometry and flow are associated with each plan, as well as any orphaned (unused) plans, flows, or  geometries.  There is a sample excel output file included.

As of 9/30/2019, quasi-unsteady flows, sediment plans, and water quality plans are not supported

 HOW TO USE:
 - Select your RAS project:
 -  Double click on this script and navigate to a .prj file OR
 -  Change the "fullpath" variable below.  If the .prj filename matches the  enclosing folder, you may provide the path of the enclosing folder. If the .prj has a different name from the enclosing folder, you must  provide the full path to the prj.
 - The default python installation for many engineers at some organizations  is the one which comes with ArcGIS.  Unfortunately, this sometimes does not have xslsxwriter available.  If your python installation has the xlsxwriter module, an excel file will be created in the RAS project directory with a table summarizing  the plans, geometries and flows.  If xlsxwriter is not available, a csv file will be created instead.

 Examples for manual entry of path name when running from an IDE:
   
 If the project file is `"D:\STYX_FLOODING\STYX_2D\STYX_2D.prj"` then either of these will work:

`fullpath = "D:\STYX_FLOODING\STYX_2D\"`

`fullpath = "D:\STYX_FLOODING\STYX_2D\STYX_2D.prj"`

 For the project file below, the file name does not match the enclosing directory, so the full path is required.

`fullpath = "D:\STYX_FLOODING\STYX_2D\AsphodelLevees.prj"`

UPDATED Jun 2022:  
 - Tweaks made to allow running in a Python3.10 environment.
 - The .bat files were useful in remapping pythong on a USACE machine.  I no longer have access to that configuration, so I can't test whether they still work.  They are not necessary on a machine with properly configured python environment variables.