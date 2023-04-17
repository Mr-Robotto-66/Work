# =========================================================================================================
# Script Name
# =========================================================================================================
# Date              : 2016-05-24
#
# Author            : Erin Philip (BCTS Okanagan - Columbia erin.n.philip@goc.bc.ca)
#
# Purpose           : This script is used to assist with Route Card Analysis, it is used in conjuction
#                     with a created MXD in the GEO environment. It looks at the block of interest and
#                     its spatial relationship with other layers that are outlined in the Look Up Table.
#
# Arguements        : This script must be run through the Route Card Toolbox with an open mxd.
#
#
# History           :   (2016-05-24) - Erin Philip
#                                      Created. Runs inside GEO. Use of Tags in map document to determine UBI
#
#                   :   (2017-03-29) - Erin Philip
#                                      Updated to run outside of GEO
#                                      Updated to run on roads and blocks based on a selection,
#                                      it will only allow a single feature to run through tool
# Chagnes
# - Added a road functionaility
# - Moved all scripting to single script
# - Moved out of GEO
# - Script is now run by a selection rather than GEO, no more than one feature can be run at a time.
# - Added a dynamic legend rather than specifying in the tables and script
# - Added new supporting data items to the script controls (All block or road specific inputs)
#
# June 6, 2017
# - Updated the 'Contained' module outside of the loop where it removes items based on the 'nooverlaplayers'
# - Contained Logic changed to be reverse. In the Module it automatically gives it a 'Y' until the block or road
#   is contained by something then it is given a 'N'
#
# November 10, 2017
#  - There were errors showing up that stated a layer did not accept "Transparency" as a layer function.
#
# December 12, 2017
# - Added Error Checking
#     ??? Added a list of approved staff to run the Route Card, issues in house of staff getting 'non trained' persons to run tool.
#         ?? Will fail if staff member IDIR is not added to the list in lower case. Can add staff at any time. Just forces staff to get training before running the tool.
#     ??? Added a check to see if there are any broken layers, will cause it to break right away rather than later on.
#
# February 8, 2018
# Changes Made:
#     - Checks to ensure that the layers are in the correct projection, it will not work properly if the layers were not in NAD 83.
#          This was added so the management of multiple layers does not need to occur.
#
# March 13, 2018
# Changes Made:
#     - Removed the changes that were made Feb 8, 2018. These caused locking issues as well as writting issues when staff did not have
#       write access to the script folder. Was rewritten with a arcpy.env.outputCoordinateSystem set to NAD 83. Also added sections to
#       reproject and select layers
#
# March 14, 2018
# Changes Made:
#     - Updated the layout function to only show layers that 'valid' by either overlapping the Shape of Interest, within a distanct of the
#       shape of interest. These layers will be turned on.
#     - Updated the SelectLayers and the Layout Functions to allow the mxd to have grouped layers. The layer names within the LUT_Processing
#       table does not change but the layers within the mxd can now be grouped. The script moves through the grouped layers.
#
#
# December 10, 2019
# Changes Made:
#     - Changed the excel output. Added a Sensitive Information field in the Processor table. All values marked as Sensitive will be added to
#          a separate excel output.
# default_to_yes
#
# =========================================================================================================

# =========================================================================================================
# =========================================================================================================
#  Process Status and Return Codes:
#      -1   - Unexpected error
#       0   - Completed Successful Run
#       1   - In progress
#       100 - Invalid Arguments
#       101 - ESRI License Error
#
# =========================================================================================================

# *********************************************************************************************************
#  Program Initialization
# *********************************************************************************************************

# ---------------------------------------------------------------------------------------------------------
#  Python and System Modules
# -----------
#  Importing all modules that will be used for the script
# ---------------------------------------------------------------------------------------------------------

from __future__ import division

import re

print '-- Importing Modules'

import sys
import string
import os
import math
import stat
import time
import getpass
import xlwt
import collections
from collections import OrderedDict
from collections import defaultdict

# ---------------------------------------------------------------------------------------------------------
#  Arguement Input
# -----------
#  Pull arguements from the tool parameters window.
# ---------------------------------------------------------------------------------------------------------

print '-- Getting Run Parameters'

PYTHON_SCRIPT = sys.argv[0]  # Python script location (automatically populated)
RCTYPE = sys.argv[1]  # The Type of Route Card Parameter. Can be Block or Road. Cannot be outside these values.
POSTFLAG = sys.argv[2]  # Post flag arguement is used to label the output with a "Final" or "Preliminary" label
if POSTFLAG == 'true':
    STATUS = "Final"
else:
    STATUS = "Preliminary"
OWNER = sys.argv[3]  # Adds the name of the staff to the output, to know who ran it.

# ---------------------------------------------------------------------------------------------------------
#  Standard/Global Variables/Constants
# ---------------------------------------------------------------------------------------------------------
print '-- Setting Global Variables'

START_TIME = time.ctime(time.time())
START_TIME_SEC = time.time()
START_TIME_SQL = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
BEEP = chr(7)

# Global Variables
#   The Processing_Variables is the collection of messaging and script specific information needed
#   throughout this program.  It is constantly updated, and generally includes all paths and configuration
#   information.  For parameters that are passed in to the script, they should generally be assigned
#   to a Processing_Variable through the Initialize() routine.

Processing_Variables = {}
RunStatus = {}
dict_labels = {}
lst_constraints = ['Grizzly Bear Habitat', 'Mountain Caribou - Okanagan', 'Mule Deer UWR', 'Bighorn Sheep Areas',
                   'Derenzy Sheep Areas', 'Mtn Goat Habitat', 'Community Watersheds', 'BEC:PPxh', 'BEC:IDFxh']

access_label = 'Access Issues (page 6 - 17)'

# ---------------------------------------------------------------------------------------------------------
#  Set up the ArcGIS Geoprocessing Environment
# ---------------------------------------------------------------------------------------------------------

import arcpy
from arcpy import env
from datetime import datetime as dt

mxd = arcpy.mapping.MapDocument("CURRENT")  # locate the mxd to work in. (set to be the mxd that is open)
df = arcpy.mapping.ListDataFrames(mxd, '*')[0]

# Set the workspace to be in memory so no extra data is created, This was changed in February 2018 due to
# projection tool not able to write to a memory database. Any data that is projected will be created in the scratch
# folder in the Scratch.gdb.
arcpy.env.workspace = 'in_memory'
# All outputs will be overwritten. This is important for any excel file that has comments written in it.
# have asked the TOC planners to move to a permanent location when comments are added to avoid them being
# overwritten.
arcpy.env.overwriteOutput = True
# Set the output coordinate system to NAD 83 so any geoprocessing that occurs the output is always in NAD 83.
arcpy.env.outputCoordinateSystem = arcpy.env.outputCoordinateSystem = arcpy.SpatialReference(3005)

# ---------------------------------------------------------------------------------------------------------
#  Inital Error Handling
# ---------------------------------------------------------------------------------------------------------

# check all layers to see if there are any broken layers, this avoids the tool breaking as it gets furthur into analysis.
brknList = [brkn.name for brkn in arcpy.mapping.ListBrokenDataSources(mxd)]
if len(brknList) > 0:
    arcpy.AddWarning("The following layers are broken and will not be included in the analysis: {} \n"
                     "Please check your permissions to ensure access to all layers".format(brknList))


# *********************************************************************************************************
#  Routines
# *********************************************************************************************************
# ---------------------------------------------------------------------------------------------------------
# !!!!!THIS IS FOR MESSAGE OUTPUTS, DO NOT CHANGE ANY OF THIS !!!!!
#
# WriteOutputToScreen(Output_Comment, Header_Style):
#
#   Prints the Output_Comment to the screen with the appropriate formatting
#       0 : "******" Used for Error Sections
#       1 : "======" Used for the Program's Header and Footer
#       2 : "------" Used for New Sections
#       3+:          No delimitation
#
# ---------------------------------------------------------------------------------------------------------
def WriteOutputToScreen(Output_Comment, Header_Style):
    if Header_Style == 0:
        # Errors get beeps, too!
        print BEEP * 3 + "\n" * 2 + "*" * 79 + "\n" * 2
    if Header_Style == 2:
        print "\n" + "-" * 79 + "\n"
    if Header_Style == 1:
        print "\n" + "=" * 79 + "\n"

    # print Output_Comment
    arcpy.AddMessage(Output_Comment)

    if Header_Style == 0:
        print "\n" * 2 + "*" * 79 + "\n" * 2
    if Header_Style == 1:
        print "\n" + "=" * 79 + "\n"


# ---------------------------------------------------------------------------------------------------------
# !!!!!THIS IS FOR MESSAGE OUTPUTS, DO NOT CHANGE ANY OF THIS !!!!!
#
#       Test if we need to do stuff, and do it if necessary.
#          Return 0 - nothing to do
#                 1 - successful
#                -1 - unknown failure
#
# ---------------------------------------------------------------------------------------------------------
def Message(Message, MessageLevel):
    # Based on the logging level, add messages to the message queues.
    #   Text Run Log   : Implimented
    #   Console        : Implimented
    #   GeoProcessor   : Implimented
    #   Processing Log : Implimented (with e-mail)
    #   SQL Database   : Not Implimented ####
    #   FGDB           : Implimented

    #   0    : Fatal Errors
    #   1    : Top Level Program
    #   2    : Main routines and events
    #   3    : Warnings
    #   4    : General Information
    #   5    : Debug level detail

    # # Send to the screen
    # if MessageLevel <= Processing_Variables['Log_Level']:
    #     WriteOutputToScreen(Message, MessageLevel)

    # Send to the GeoProcessor Message Queue
    #### At ArcGIS 9.3 SP1 there is a bug that also causes any GeoProcessor messages to be echoed to the console.
    #    Added the ToolBoxRun parameter to allow the script to add messages to the GeoProcessor Messaging properly
    if MessageLevel < 4 and MessageLevel <= Processing_Variables['Log_Level'] and Processing_Variables[
        'ToolBoxRun'] == 'TRUE':
        if MessageLevel == 0:
            arcpy.AddError(Message)
        if MessageLevel == 1 or MessageLevel == 2:
            arcpy.AddMessage(Message)
        if MessageLevel == 3:
            arcpy.AddWarning(Message)

    # Send to the FGDB Processing Log
    if MessageLevel < 4 and MessageLevel <= Processing_Variables[
        'Log_Level'] and 'Log_FGDB_Table' in Processing_Variables:
        rows = arcpy.InsertCursor(Processing_Variables['Log_FGDB_Table'])

        row = rows.NewRow()
        row.SetValue("Script_Name", PYTHON_SCRIPT)
        row.SetValue("Event_Time", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        row.SetValue("Message_Level", MessageLevel)
        row.SetValue("Message", Message)
        row.SetValue("Start_Time", START_TIME_SQL)

        rows.InsertRow(row)
        del row
        del rows

    return (1)


# =========================================================================================================
#  Standard Routines
# =========================================================================================================
# ---------------------------------------------------------------------------------------------------------
# Initialize()
#
#       Set up all of the paths and variables that we need to run this script.
#
# ---------------------------------------------------------------------------------------------------------
def Initialize():
    print '    - Initializing Script'

    Processing_Variables['ToolBoxRun'] = 'TRUE'
    Processing_Variables['Log_Level'] = 4
    Processing_Variables['PROCESS_INFO'] = "RUNNING " + PYTHON_SCRIPT
    Processing_Variables['PROCESS_STATUS'] = 1
    Processing_Variables['PROCESS_LOG'] = []

    Message('Executing: ' + PYTHON_SCRIPT, 1)
    Message('    Start Time: ' + START_TIME, 2)
    Message('    Python ' + sys.version, 4)

    # Variables for all the paths to directories and gdbs. This is where global variables for these items should be saved.
    Processing_Variables['Script_Directory_Path'] = os.path.split(PYTHON_SCRIPT)[0]
    Processing_Variables['Script_Name'] = os.path.split(PYTHON_SCRIPT)[1]
    Processing_Variables['Log_Directory_Path'] = Processing_Variables['Script_Directory_Path'] + r'\log'
    Processing_Variables['Project_Base_Path'] = os.path.split(Processing_Variables['Script_Directory_Path'])[0]
    Processing_Variables['Supporting_Data_Directory_Path'] = Processing_Variables[
                                                                 'Script_Directory_Path'] + r'\Supporting Data'
    Processing_Variables['Supporting_Data_GDB'] = Processing_Variables[
                                                      'Supporting_Data_Directory_Path'] + r'\SupportingData.gdb'
    Processing_Variables['Scratch_Data_Directory_Path'] = Processing_Variables['Script_Directory_Path'] + r'\Scratch'
    Processing_Variables['Scratch_Data_GDB'] = Processing_Variables['Scratch_Data_Directory_Path'] + r'\Scratch.gdb'

    # -----------------------------------------------------------------------------------------------------
    #  Script Controls
    # -----------------------------------------------------------------------------------------------------
    # This section creates a list of all the controls found in the LUT_ScriptControls table and converts them all to variables.

    Processing_Variables['LUT_ScriptControls'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_ScriptControls'

    Processing_Variables['Variables'] = {}
    for row in arcpy.da.SearchCursor(Processing_Variables['LUT_ScriptControls'], ['Script_Variable', 'Variable_Value']):
        Processing_Variables['Variables'][row[0]] = row[1]

    initialize_variables()

    # -----------------------------------------------------------------------------------------------------
    #  Script Specific Variables/Constants
    # -----------------------------------------------------------------------------------------------------
    # String Variables
    # Processing_Variables['UNIQUEID'] = ''
    # Processing_Variables['Status'] = STATUS
    # Processing_Variables['Owner'] = OWNER
    # Processing_Variables['RCTYPE'] = RCTYPE
    # Processing_Variables['SearchBuffer'] = ''
    # Processing_Variables['ShapeOfInterest'] = ''
    #
    # Processing_Variables['OpArea'] = ''
    # Processing_Variables['LandscapeUnit'] = ''
    # Processing_Variables['LegalLocation'] = ''
    # Processing_Variables['PolygonSelection'] = ''
    # Processing_Variables['UniqueIDField'] = ''
    # Processing_Variables['TitleFields'] = ''
    # Processing_Variables['ExtentString'] = ''
    #
    # # Numeric Variables
    # Processing_Variables['newXMax'] = 0
    # Processing_Variables['newXMin'] = 0
    # Processing_Variables['newYMax'] = 0
    # Processing_Variables['newYMin'] = 0
    # Processing_Variables['oldExtent'] = None
    #
    # # Path Variables
    # Processing_Variables['ProcessingLookupTable'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_Processing'
    # Processing_Variables['LegendLookupTable'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_Legend'
    # Processing_Variables['LegalArea'] = Processing_Variables['Supporting_Data_GDB'] + '\\LegalAreas'
    # Processing_Variables['OutputLocation'] = Processing_Variables['Variables']['OutputLocation']
    # Processing_Variables['InputFeatureClass'] = r'in_memory\input_fc'
    #
    # # Lists
    # Processing_Variables['ShapeArea'] = []
    # Processing_Variables['SearchBufferArea'] = []
    # Processing_Variables['ExistingBEC'] = []
    # Processing_Variables['whacodes'] = []
    # Processing_Variables['ungulatecodes'] = []
    # Processing_Variables['bands'] = []
    # Processing_Variables['NoOverlap'] = []
    # Processing_Variables['LayerList'] = []
    # Processing_Variables['Sensitive_Information'] = []
    #
    # # Dictionaries
    # Processing_Variables['Applicable'] = defaultdict(str)
    # Processing_Variables['SelectionLayers'] = defaultdict(str)
    # Processing_Variables['AttributeInformation'] = defaultdict(list)
    # Processing_Variables['CannedStatement'] = defaultdict(str)
    # Processing_Variables['layerdict'] = defaultdict(str)
    # Processing_Variables['Title'] = defaultdict(str)

    return (1)


def initialize_variables():
    Processing_Variables['UNIQUEID'] = ''
    Processing_Variables['Status'] = STATUS
    Processing_Variables['Owner'] = OWNER
    Processing_Variables['RCTYPE'] = RCTYPE
    Processing_Variables['SearchBuffer'] = ''
    Processing_Variables['ShapeOfInterest'] = ''

    Processing_Variables['OpArea'] = ''
    Processing_Variables['LandscapeUnit'] = ''
    Processing_Variables['LegalLocation'] = ''
    Processing_Variables['PolygonSelection'] = ''
    Processing_Variables['UniqueIDField'] = ''
    Processing_Variables['TitleFields'] = ''
    Processing_Variables['ExtentString'] = ''

    # Numeric Variables
    Processing_Variables['newXMax'] = 0
    Processing_Variables['newXMin'] = 0
    Processing_Variables['newYMax'] = 0
    Processing_Variables['newYMin'] = 0
    Processing_Variables['oldExtent'] = None

    # Path Variables
    Processing_Variables['ProcessingLookupTable'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_Processing'
    Processing_Variables['LegendLookupTable'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_Legend'
    Processing_Variables['LegalArea'] = Processing_Variables['Supporting_Data_GDB'] + '\\LegalAreas'
    Processing_Variables['OutputLocation'] = Processing_Variables['Variables']['OutputLocation']
    Processing_Variables['InputFeatureClass'] = r'in_memory\input_fc'

    # Lists
    Processing_Variables['ShapeArea'] = []
    Processing_Variables['SearchBufferArea'] = []
    Processing_Variables['ExistingBEC'] = []
    Processing_Variables['whacodes'] = []
    Processing_Variables['ungulatecodes'] = []
    Processing_Variables['bands'] = []
    Processing_Variables['NoOverlap'] = []
    Processing_Variables['LayerList'] = []
    Processing_Variables['Sensitive_Information'] = []

    # Dictionaries
    Processing_Variables['Applicable'] = defaultdict(str)
    Processing_Variables['SelectionLayers'] = defaultdict(str)
    Processing_Variables['AttributeInformation'] = defaultdict(list)
    Processing_Variables['CannedStatement'] = defaultdict(str)
    Processing_Variables['layerdict'] = defaultdict(str)
    Processing_Variables['Title'] = defaultdict(str)

# ---------------------------------------------------------------------------------------------------------
# Reset MXD()
#
#
# ---------------------------------------------------------------------------------------------------------
def ResetMXD():
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "Reset MXD"
    Message('    - Reset MXD', 2)

    arcpy.env.workspace = 'in_memory'
    if Processing_Variables['oldExtent']:
        arcpy.env.extent = Processing_Variables['oldExtent']

    # Run through the layers in the Table of Contents and remove that are 'lingering' from previous runs of this tool.
    for layer in arcpy.mapping.ListLayers(mxd, ):
        if layer.name in ["ShapeOfInterest", "SearchBuffer", "Road of Interest", "Block of Interest"]:
            arcpy.mapping.RemoveLayer(df, layer)
        if "Meter Search Buffer" in layer.name:
            arcpy.mapping.RemoveLayer(df, layer)

    return (Processing_Variables['PROCESS_STATUS'])


# ---------------------------------------------------------------------------------------------------------
# SelectBlocks()
#
#
# ---------------------------------------------------------------------------------------------------------
def SelectBlocks():
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "SelectBlocks"

    Message('    - Selecting Blocks', 2)

    # # Test to see if a SINGLE block selection exists.
    # desc = arcpy.Describe(Processing_Variables['Variables']['InputBlockFeatureClass'])
    # # There is more than one block selected - tool cannot keep going
    # if len(desc.FIDSet.split(';')) > 1:
    #     Message("   - Error with Blocks Selection - Stopping Tool", 0)
    #     sys.exit(100)
    # # There are no blocks selected - Tool cannot keep going.
    # elif str(desc.FIDSet) == "" :
    #     Message("   - No Block Selected - Stopping Tool", 0)
    #     sys.exit(100)

    # Create a single block featureclass in memory to be used through the script.
    arcpy.Select_analysis(Processing_Variables['InputFeatureClass'], r'in_memory\ShapeOfInterest')

    Processing_Variables['ShapeOfInterest'] = r'in_memory\ShapeOfInterest'
    Processing_Variables['UniqueIDField'] = Processing_Variables['Variables']['InputBlockField']

    # Get the attribute information associated with this block as it was outlined in teh LUT_Script Controls Table.
    # This information will be used for headers and saving the file.
    cursorfieldlist = Processing_Variables['Variables']['InputBlockTitleFields'].split(',')
    for row in arcpy.da.SearchCursor(Processing_Variables['ShapeOfInterest'], cursorfieldlist):
        for item in cursorfieldlist:
            Processing_Variables['Title'][item] = row[cursorfieldlist.index(item)]
    with arcpy.da.SearchCursor(Processing_Variables['ShapeOfInterest'], Processing_Variables['UniqueIDField']) as s_cursor:
        for row in s_cursor:
            Processing_Variables['UNIQUEID'] = int(row[0])

    Processing_Variables['TitleFields'] = Processing_Variables['Variables']['InputBlockTitleFields']

    return (Processing_Variables['PROCESS_STATUS'])


# ---------------------------------------------------------------------------------------------------------
# SelectRoads()
#
#
# ---------------------------------------------------------------------------------------------------------
def SetupRoads():
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "Select Roads"

    Message('    - Selecting Roads', 2)

    # desc = arcpy.Describe(Processing_Variables['Variables']['InputRoadFeatureClass'])
    # if len(desc.FIDSet.split(';')) > 1:
    #     Message("   - Error with Roads Selection - Stopping Tool",0)
    #     sys.exit(100)
    # elif str(desc.FIDSet) == "" :
    #     Message("   - No Road Selected - Stopping Tool",0)
    #     sys.exit(100)
    # Create a feature class of selected road feature
    arcpy.CopyFeatures_management(Processing_Variables['InputFeatureClass'], r'in_memory\RoadSelect')

    # buffer all roads into
    arcpy.Buffer_analysis(r'in_memory\RoadSelect', r"in_memory\ShapeOfInterest", '10 Meters', dissolve_option='ALL')

    Processing_Variables['ShapeOfInterest'] = r'in_memory\ShapeOfInterest'
    Processing_Variables['UniqueIDField'] = Processing_Variables['Variables']['InputRoadField']

    cursorfieldlist = Processing_Variables['Variables']['InputRoadTitleFields'].split(',')
    for row in arcpy.da.SearchCursor(r'in_memory\RoadSelect', cursorfieldlist):
        for item in cursorfieldlist:
            Processing_Variables['Title'][item] = row[cursorfieldlist.index(item)]

    Processing_Variables['TitleFields'] = Processing_Variables['Variables']['InputRoadTitleFields']

    return (Processing_Variables['PROCESS_STATUS'])
    Message('           RoadSelect is' + r'in_memory\RoadSelect')


# ---------------------------------------------------------------------------------------------------------
# PrepMXD()
#
#
# ---------------------------------------------------------------------------------------------------------
def PrepMXD():
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "Prepping MXD"

    Message('    - Prepping MXD', 2)

    # Create the Search Buffer, add to map and apply symbology from a premade lyr file.
    BufferDistance = Processing_Variables['Variables']['SearchBufferDistance'] + ' Meters'
    arcpy.Buffer_analysis(r'in_memory\ShapeOfInterest', r"in_memory\SearchBuffer", BufferDistance)
    Processing_Variables['SearchBuffer'] = r"in_memory\SearchBuffer"

    addbuffer = arcpy.mapping.Layer(Processing_Variables['SearchBuffer'])
    arcpy.mapping.AddLayer(df, addbuffer)

    arcpy.ApplySymbologyFromLayer_management("SearchBuffer", Processing_Variables[
        'Supporting_Data_Directory_Path'] + r'\SearchBuffer.lyr')

    addlayer = arcpy.mapping.Layer(Processing_Variables['ShapeOfInterest'])
    arcpy.mapping.AddLayer(df, addlayer)
    arcpy.ApplySymbologyFromLayer_management("ShapeOfInterest", Processing_Variables[
        'Supporting_Data_Directory_Path'] + r'\ShapeOfInterest.lyr')

    del Processing_Variables['ShapeArea'][:]
    del Processing_Variables['SearchBufferArea'][:]
    del Processing_Variables['ExistingBEC'][:]

    # Grab the spatial shapes from the Search Buffer and Shape of Interest for future analysis
    with arcpy.da.SearchCursor(Processing_Variables['ShapeOfInterest'], ["SHAPE@"]) as cursor:
        for row in cursor:
            Processing_Variables['ShapeArea'].append(row[0])

    with arcpy.da.SearchCursor(Processing_Variables['SearchBuffer'], ["SHAPE@"]) as cursor:
        for row in cursor:
            Processing_Variables['SearchBufferArea'].append(row[0])

    # Turn off all layers except the UBI of interest Layer
    for layer in arcpy.mapping.ListLayers(mxd):
        if layer.name in ["ShapeOfInterest", 'SearchBufferDistance']:
            pass
        else:
            layer.visible = False

    # Set the extent in which you set the Processing Extent
    ext = addbuffer.getExtent()
    df.extent = ext
    if df.scale < int(Processing_Variables['Variables']['DataframeScale']):
        df.scale = int(Processing_Variables['Variables']['DataframeScale'])
    else:
        df.scale = round(df.scale, -3) + 1000

    Processing_Variables['newXMax'] = ext.XMax + int(Processing_Variables['Variables']['ExtentDistance'])
    Processing_Variables['newXMin'] = ext.XMin - int(Processing_Variables['Variables']['ExtentDistance'])
    Processing_Variables['newYMax'] = ext.YMax + int(Processing_Variables['Variables']['ExtentDistance'])
    Processing_Variables['newYMin'] = ext.YMin - int(Processing_Variables['Variables']['ExtentDistance'])

    Processing_Variables['oldExtent'] = arcpy.env.extent

    # Set Processing Extent
    arcpy.env.extent = arcpy.Extent(Processing_Variables['newXMin'], Processing_Variables['newYMin'],
                                    Processing_Variables['newXMax'], Processing_Variables['newYMax'])

    return (Processing_Variables['PROCESS_STATUS'])


# ---------------------------------------------------------------------------------------------------------
# DetermineRequirements()
#
#
# ---------------------------------------------------------------------------------------------------------
def DetermineRequirements():
    # find what legal values are we constrained by.
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "PreppingMXD"

    Message('    - Determining What Card to Run', 2)

    for row in arcpy.da.SearchCursor(Processing_Variables['LegalArea'], ['LegalArea', "SHAPE@"]):
        for shape in Processing_Variables['ShapeArea']:
            if shape.within(row[1]) or row[1].overlaps(shape):
                Processing_Variables['LegalLocation'] = row[0]

    return (Processing_Variables['PROCESS_STATUS'])


# ---------------------------------------------------------------------------------------------------------
# SelectLayers()
#
#       Select all layers that exist in the Processing Lookup table to make processing faster. Write to memory
#
# ---------------------------------------------------------------------------------------------------------
def SelectLayers():
    Processing_Variables['LayerList'] = []
    # Processing_Variables['SelectionLayers'].clear()
    # Processing_Variables['NoOverlap'] = []
    # Create a selection of all layers within the LUT Processing Table to make processing faster
    Message('    - Analyzing Layers in the Defined Extent', 2)
    for row in arcpy.da.SearchCursor(Processing_Variables['ProcessingLookupTable'],
                                     ['Layer_List', Processing_Variables['LegalLocation']]):
        if row[1] > 0:
            for item in str(row[0]).replace("None", '').split(';'):
                try:
                    item = item.split(':')[0]
                except:
                    item = item
                if item not in brknList:
                    Processing_Variables['LayerList'].append(item)

    layercount = 0
    for layer in arcpy.mapping.ListLayers(mxd):
        lyrname = layer.name
        arcpy.AddMessage (lyrname)
        if lyrname not in Processing_Variables['LayerList'] or layer.isGroupLayer:
            pass
        else:
            layercount += 1
            # test to see if the layer is in a NAD 83 projection to ensure the spatial geometry is handled properly in the processes later.
            desc = arcpy.Describe(layer)
            sellayer = layer
            projection = desc.spatialReference.name

            count = arcpy.GetCount_management(sellayer)
            Message('       - ' + lyrname + ": " + str(count), 4)

            if str(count) == '0' and projection in ["NAD_1983_BC_Environment_Albers"] and lyrname <> "BCTS SU":
                Processing_Variables['NoOverlap'].append(layer.name)
            else:
                if projection not in ["NAD_1983_BC_Environment_Albers"]:
                    arcpy.env.extent = 'MAXOF'
                    arcpy.Select_analysis(sellayer, 'in_memory\\tempSelection')

                    arcpy.env.extent = arcpy.Extent(Processing_Variables['newXMin'], Processing_Variables['newYMin'],
                                                    Processing_Variables['newXMax'], Processing_Variables['newYMax'])
                    arcpy.Select_analysis('in_memory\\tempSelection', 'in_memory\\Selection_' + str(layercount))

                    count = arcpy.GetCount_management('in_memory\\Selection_' + str(layercount))
                    Message('           -Projected Count: ' + str(count), 4)
                    if str(count) == '0':
                        Processing_Variables['NoOverlap'].append(layer.name)

                else:
                    arcpy.Select_analysis(sellayer, 'in_memory\\Selection_' + str(layercount))

                oldname = str(layer.name)
                newname = 'Selection_' + str(layercount)
                Processing_Variables['SelectionLayers'][oldname] = newname


# ----------------------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------
def containsoverlap(layer, label, cursorfieldlist, Buffer_Meters, def_query, layer_name, num_layers):
    lst_attributes = []
    lst_habitat = []
    lst_shapes = []
    canned_statement = Processing_Variables['CannedStatement'][label]

    with arcpy.da.SearchCursor(layer, cursorfieldlist, where_clause=def_query) as cursor:
        for row in cursor:
            if str(row[0]) == "None":
                pass
            else:
                for area in Processing_Variables['ShapeArea']:
                    if not row[0].disjoint(area):
                        attribute = ''
                        Processing_Variables['Applicable'][label] = 'Y'
                        if label not in ['Lakeshore Management Zones']:
                            if len(cursorfieldlist) == 2:
                                if canned_statement:
                                    if '[1]' in canned_statement:
                                        if row[1] not in lst_habitat:
                                            lst_habitat.append(row[1])
                                        # canned_statement = canned_statement.replace('[1]', row[1])
                                        # Processing_Variables['CannedStatement'][label] = canned_statement
                                        continue
                                attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                            elif len(cursorfieldlist) > 2:
                                for num in range(1, len(cursorfieldlist)):
                                    if num == 2:
                                        attribute += ' ('
                                    attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                                    if 2 <= num < (len(cursorfieldlist) - 1):
                                        attribute += ', '
                                    elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                        attribute += ')'

                            if attribute not in lst_attributes and attribute != '':
                                lst_attributes.append(attribute)

                        else:
                            if len(cursorfieldlist) > 1:
                                for num in range(1, len(cursorfieldlist)):
                                    str_label = dict_labels[cursorfieldlist[num]]
                                    attribute = '{} = {}'.format(str_label, row[num])

                                    if attribute not in lst_attributes and attribute != '':
                                        lst_attributes.append(attribute)

                        lst_shapes.append(row[0])
    concat_attribute = ''
    if label == 'Visual Quality Objectives':
        concat_attribute += 'VQO = '
    elif label == 'Migratory Birds':
        concat_attribute += 'Overlaps Habitat Rank:'

    if label.startswith('Old Growth Deferral') and Processing_Variables['Applicable'][label] == 'Y':
        Processing_Variables['CannedStatement'][label] = 'TAP field verification and rationale required.  {0}'\
            .format(Processing_Variables['CannedStatement'][label])

    if len(lst_attributes) > 0:
        if num_layers > 1:
            concat_attribute += '{0}: {1}'.format(layer_name, ', '.join(lst_attributes))
        else:
            concat_attribute += ', '.join(lst_attributes)
    # if len(lst_attributes) > 0:
    #     delim = ', '
    #     # if len(cursorfieldlist) == 2:
    #     #     delim = ', '
    #     # elif len(cursorfieldlist) > 2:
    #     #     delim = '\n'
    #
    #     for attr in lst_attributes:
    #         concat_attribute += attr
    #         if attr != lst_attributes[-1]:
    #             concat_attribute += delim

    if len(lst_habitat) > 0:
        str_habitat = ''
        for i in range(0, len(lst_habitat)):
            str_habitat += lst_habitat[i]
            if len(lst_habitat) > 1:
                if i == len(lst_habitat) - 2:
                    str_habitat += ' and '
                elif i < len(lst_habitat) - 2:
                    str_habitat += ', '

        canned_statement = canned_statement.replace('[1]', str_habitat)
        Processing_Variables['CannedStatement'][label] = canned_statement


    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    if label in ['Community Watersheds', 'Fisheries Sensitive Watershed'] and concat_attribute != '':
        if not Processing_Variables['AttributeInformation']['Hydrological']:
            Processing_Variables['AttributeInformation']['Hydrological'].append('Located within ' + concat_attribute)
        else:
            Processing_Variables['AttributeInformation']['Hydrological'][0] += ', ' + concat_attribute
        # try:
        #
        #     Processing_Variables['AttributeInformation']['Hydrological'].append('Located within ' + concat_attribute)
        # except:
        #     Processing_Variables['AttributeInformation']['Hydrological'] = ['Located within ' + concat_attribute]
    if label == 'Karst Potential':
        for b in [100, 200, 500]:
            check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=b,
                        layer_name=layer_name, lst_shapes=lst_shapes)

    else:
        check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=Buffer_Meters,
                    layer_name=layer_name, lst_shapes=lst_shapes)

    # if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
    #     with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
    #         for row in cursor:
    #             for area in Processing_Variables['ShapeArea']:
    #                 buff = area.buffer(Buffer_Meters)
    #                 if not row[0].disjoint(buff):
    #                     Processing_Variables['Applicable'][label] = 'X -' + str(Buffer_Meters)
    #                     break
    #
    # else:
    #     pass

    return


# ----------------------------------------------------------------------------------------------------
#
def SOMC(layer, label, cursorfieldlist, Buffer_Meters, def_query, layer_name, num_layers):
    lst_attributes = []
    lst_habitat = []
    lst_shapes = []
    canned_statement = Processing_Variables['CannedStatement'][label]

    with arcpy.da.SearchCursor(layer, cursorfieldlist, where_clause=def_query) as cursor:
        for row in cursor:
            if str(row[0]) == "None":
                pass
            else:
                for area in Processing_Variables['ShapeArea']:
                    if not row[0].disjoint(area):
                        attribute = ''
                        Processing_Variables['Applicable'][label] = 'SOMC'
                        if label not in ['Lakeshore Management Zones']:
                            if len(cursorfieldlist) == 2:
                                if canned_statement:
                                    if '[1]' in canned_statement:
                                        if row[1] not in lst_habitat:
                                            lst_habitat.append(row[1])
                                        # canned_statement = canned_statement.replace('[1]', row[1])
                                        # Processing_Variables['CannedStatement'][label] = canned_statement
                                        continue
                                attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                            elif len(cursorfieldlist) > 2:
                                for num in range(1, len(cursorfieldlist)):
                                    if num == 2:
                                        attribute += ' ('
                                    attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                                    if 2 <= num < (len(cursorfieldlist) - 1):
                                        attribute += ', '
                                    elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                        attribute += ')'

                            if attribute not in lst_attributes and attribute != '':
                                lst_attributes.append(attribute)

                        else:
                            if len(cursorfieldlist) > 1:
                                for num in range(1, len(cursorfieldlist)):
                                    str_label = dict_labels[cursorfieldlist[num]]
                                    attribute = '{} = {}'.format(str_label, row[num])

                                    if attribute not in lst_attributes and attribute != '':
                                        lst_attributes.append(attribute)

                        lst_shapes.append(row[0])
    concat_attribute = ''
    if label == 'Visual Quality Objectives':
        concat_attribute += 'VQO = '
    elif label == 'Migratory Birds':
        concat_attribute += 'Overlaps Habitat Rank:'

    if label.startswith('Old Growth Deferral') and Processing_Variables['Applicable'][label] == 'Y':
        Processing_Variables['CannedStatement'][label] = 'TAP field verification and rationale required.  {0}'\
            .format(Processing_Variables['CannedStatement'][label])

    if len(lst_attributes) > 0:
        if num_layers > 1:
            concat_attribute += '{0}: {1}'.format(layer_name, ', '.join(lst_attributes))
        else:
            concat_attribute += ', '.join(lst_attributes)
    # if len(lst_attributes) > 0:
    #     delim = ', '
    #     # if len(cursorfieldlist) == 2:
    #     #     delim = ', '
    #     # elif len(cursorfieldlist) > 2:
    #     #     delim = '\n'
    #
    #     for attr in lst_attributes:
    #         concat_attribute += attr
    #         if attr != lst_attributes[-1]:
    #             concat_attribute += delim

    if len(lst_habitat) > 0:
        str_habitat = ''
        for i in range(0, len(lst_habitat)):
            str_habitat += lst_habitat[i]
            if len(lst_habitat) > 1:
                if i == len(lst_habitat) - 2:
                    str_habitat += ' and '
                elif i < len(lst_habitat) - 2:
                    str_habitat += ', '

        canned_statement = canned_statement.replace('[1]', str_habitat)
        Processing_Variables['CannedStatement'][label] = canned_statement


    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    if label in ['Community Watersheds', 'Fisheries Sensitive Watershed'] and concat_attribute != '':
        if not Processing_Variables['AttributeInformation']['Hydrological']:
            Processing_Variables['AttributeInformation']['Hydrological'].append('Located within ' + concat_attribute)
        else:
            Processing_Variables['AttributeInformation']['Hydrological'][0] += ', ' + concat_attribute
        # try:
        #
        #     Processing_Variables['AttributeInformation']['Hydrological'].append('Located within ' + concat_attribute)
        # except:
        #     Processing_Variables['AttributeInformation']['Hydrological'] = ['Located within ' + concat_attribute]
    if label == 'Karst Potential':
        for b in [100, 200, 500]:
            check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=b,
                        layer_name=layer_name, lst_shapes=lst_shapes)

    else:
        check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=Buffer_Meters,
                    layer_name=layer_name, lst_shapes=lst_shapes)

    # if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
    #     with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
    #         for row in cursor:
    #             for area in Processing_Variables['ShapeArea']:
    #                 buff = area.buffer(Buffer_Meters)
    #                 if not row[0].disjoint(buff):
    #                     Processing_Variables['Applicable'][label] = 'X -' + str(Buffer_Meters)
    #                     break
    #
    # else:
    #     pass

    return
#
# ----------------------------------------------------------------------------------------------------
def overlaptouching(layer, label, cursorfieldlist, Buffer_Meters):
    lst_attributes = []

    with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
        for row in cursor:
            if str(row[0]) == "None":
                pass
            else:
                for area in Processing_Variables['ShapeArea']:
                    buff = area.buffer(Buffer_Meters)
                    if not row[0].disjoint(area) or not row[0].disjoint(buff):
                        attribute = ''
                        Processing_Variables['Applicable'][label] = 'Y'
                        if len(cursorfieldlist) == 2:
                            attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                        elif len(cursorfieldlist) > 2:
                            for num in range(1, len(cursorfieldlist)):
                                if num == 2:
                                    attribute += ' ('
                                attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                                if 2 <= num < (len(cursorfieldlist) - 1):
                                    attribute += ', '
                                elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                    attribute += ')'
                        if attribute not in lst_attributes:
                            lst_attributes.append(attribute)

    concat_attribute = ''
    if len(lst_attributes) > 0:
        delim = ', '
        # if len(cursorfieldlist) == 2:
        #     delim = ', '
        # elif len(cursorfieldlist) > 2:
        #     delim = '\n'

        for attr in lst_attributes:
            concat_attribute += attr
            if attr != lst_attributes[-1]:
                concat_attribute += delim

    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    # if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
    #     with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
    #         for row in cursor:
    #             for area in Processing_Variables['ShapeArea']:
    #                 buff = area.buffer(Buffer_Meters)
    #                 if not row[0].disjoint(buff):
    #                     Processing_Variables['Applicable'][label] = 'X -' + str(Buffer_Meters)
    #                     break
    # else:
    #     pass

    return


# ----------------------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------
def contained(layer, label, cursorfieldlist):
    concatattribute = ''

    # Must make it a Yes, will change to N IF they are contained in each other.
    Processing_Variables['Applicable'][label] = 'Y'
    with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
        for row in cursor:
            if str(row[0]) == "None":
                pass
            else:
                for area in Processing_Variables['ShapeArea']:
                    if row[0].contains(area):
                        Processing_Variables['Applicable'][label] = 'N'
                        break
                    else:
                        pass

    return


# ----------------------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------


def pointline(layer, label, cursorfieldlist, Buffer_Meters, layer_name, num_layers):
    lst_attributes = []
    lst_shapes = []

    with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
        for row in cursor:
            if str(row[0]) == 'None':
                pass
            else:
                for area in Processing_Variables['ShapeArea']:
                    if area.contains(row[0]) or area.touches(row[0]) or row[0].crosses(area) or not row[0].disjoint(area):
                        attribute = ''
                        Processing_Variables['Applicable'][label] = 'Y'
                        if len(cursorfieldlist) == 2:
                            attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                        elif len(cursorfieldlist) > 2:
                            for num in range(1, len(cursorfieldlist)):
                                if num == 2:
                                    attribute += ' ('
                                attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                                if 2 <= num < (len(cursorfieldlist) - 1):
                                    attribute += ', '
                                elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                    attribute += ')'
                        if attribute not in lst_attributes:
                            lst_attributes.append(attribute)
                        lst_shapes.append(row[0])

    concat_attribute = ''
    if len(lst_attributes) > 0:
        if num_layers > 1:
            concat_attribute = '{0}: {1}'.format(layer_name, ', '.join(lst_attributes))
        else:
            concat_attribute = ', '.join(lst_attributes)
    #     delim = ', '
    #     # if len(cursorfieldlist) == 2:
    #     #     delim = ', '
    #     # elif len(cursorfieldlist) > 2:
    #     #     delim = '\n'
    #
    #     for attr in lst_attributes:
    #         concat_attribute += attr
    #         if attr != lst_attributes[-1]:
    #             concat_attribute += delim

    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    if label == 'Water Purveyor':
        for b in [100, 500, 1000]:
            if layer_name == 'Points of Diversion - Lines' and b > 100:
                break
            check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=b,
                        layer_name=layer_name, lst_shapes=lst_shapes)
    else:
        check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=Buffer_Meters,
                    layer_name=layer_name, lst_shapes=lst_shapes)
    # lst_attributes = []
    # with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
    #     for row in cursor:
    #         for area in Processing_Variables['ShapeArea']:
    #             buff = area.buffer(Buffer_Meters)
    #             if not row[0].disjoint(buff) and row[0] not in lst_shapes:
    #                 attribute = ''
    #                 if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
    #                     Processing_Variables['Applicable'][label] = 'X -' + str(Buffer_Meters)
    #                 if len(cursorfieldlist) == 2:
    #                     attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
    #                 elif len(cursorfieldlist) > 2:
    #                     for num in range(1, len(cursorfieldlist)):
    #                         if num == 2:
    #                             attribute += ' ('
    #                         attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
    #                         if 2 <= num < (len(cursorfieldlist) - 1):
    #                             attribute += ', '
    #                         elif num >= 2 and num == (len(cursorfieldlist) - 1):
    #                             attribute += ')'
    #                 if attribute not in lst_attributes:
    #                     lst_attributes.append(attribute)
    #                 lst_shapes.append(row[0])
    # concat_attribute = ''
    # if len(lst_attributes) > 0:
    #     concat_attribute = 'Within {0}m of {1}: {2}'.format(int(Buffer_Meters), layer_name, ', '.join(lst_attributes))
    #
    # try:
    #     Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    # except:
    #     Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    return


def check_range(layer, label, cursorfieldlist, Buffer_Meters, layer_name, lst_shapes):
    lst_attributes = []
    with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
        for row in cursor:
            for area in Processing_Variables['ShapeArea']:
                buff = area.buffer(Buffer_Meters)
                if not row[0].disjoint(buff) and row[0] not in lst_shapes:
                    attribute = ''
                    if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y','S']:
                        Processing_Variables['Applicable'][label] = 'X -' + str(Buffer_Meters)
                    if len(cursorfieldlist) == 2:
                        attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                    elif len(cursorfieldlist) > 2:
                        for num in range(1, len(cursorfieldlist)):
                            if num == 2:
                                attribute += ' ('
                            attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                            if 2 <= num < (len(cursorfieldlist) - 1):
                                attribute += ', '
                            elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                attribute += ')'
                    if attribute not in lst_attributes:
                        lst_attributes.append(attribute)
                    lst_shapes.append(row[0])
    concat_attribute = ''
    if len(lst_attributes) > 0:
        if len(lst_attributes) == 1 and lst_attributes[0] == '':
            concat_attribute = 'Within {0}m of {1}'.format(int(Buffer_Meters), layer_name)
        else:
            concat_attribute = 'Within {0}m of {1}: {2}'.format(int(Buffer_Meters), layer_name,
                                                                ', '.join(lst_attributes))

    Processing_Variables['AttributeInformation'][label].append(concat_attribute)

    if label in ['Community Watersheds', 'Fisheries Sensitive Watershed'] and concat_attribute != '':
        Processing_Variables['AttributeInformation']['Hydrological'].append(concat_attribute)


    return lst_shapes

# ---------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------
def BECDefinedAreas(label, query, queryfield, layer):
    # Search through overlapping BEC layers to find the values that intersect with the block/road
    if len(Processing_Variables['ExistingBEC']) < 1:
        with arcpy.da.SearchCursor(layer, [queryfield, 'SHAPE@']) as cursor:
            for row in cursor:
                for area in Processing_Variables['ShapeArea']:
                    if not row[1].disjoint(area):
                        Processing_Variables['ExistingBEC'].append(row[0].replace(' ', ''))
    # The above code created a list of all bec that overlaps with the block.
    # Loop through the list and see if it falls into any of the required values in the LUT_PRocessing Table.
    # If a value has a * that means that it is looking for any BEC value that has that Zone and Subzone and Variant.

    canned_statement = str(Processing_Variables['CannedStatement'][label]).replace('None', '')

    applicable_prefix = ''
    na_prefix = ''
    if label not in ['Williamsons Sapsucker (By BEC)']:
        applicable_prefix = 'Occurs in this BEC Variant. '
        na_prefix = 'Not typically found in this BEC Variant. '
        if label not in ['Interior Western Screech-Owl']:
            applicable_prefix += 'Assess for suitable habitat: '

    for zone in query:
        if '*' in zone:
            for bec in Processing_Variables['ExistingBEC']:
                if zone.replace('*', '') in bec:
                    Processing_Variables['Applicable'][label] = 'BA'
                    if canned_statement != '':
                        Processing_Variables['CannedStatement'][label] = applicable_prefix + canned_statement
        else:
            if zone in Processing_Variables['ExistingBEC']:
                Processing_Variables['Applicable'][label] = 'BA'
                if canned_statement != '':
                    Processing_Variables['CannedStatement'][label] = applicable_prefix + canned_statement

    if Processing_Variables['Applicable'][label] != 'BA' and label not in ['Williamsons Sapsucker (By BEC)']:
        if canned_statement != '':
            Processing_Variables['Applicable'][label] = 'NBA'
            Processing_Variables['CannedStatement'][label] = na_prefix + canned_statement
    elif Processing_Variables['Applicable'][label] != 'BA' and label in ['Williamsons Sapsucker (By BEC)']:
        if canned_statement != '':
            Processing_Variables['Applicable'][label] = 'NBA'
            Processing_Variables['CannedStatement'][label] = na_prefix

    return


# ----------------------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------
def WildlifeHabitatAreas(label, query, queryfield, layer):
    lst_attributes = []

    if len(Processing_Variables['whacodes']) < 1:
        # Message("Analyzing: Wildlife Habitat Areas",4)
        with arcpy.da.SearchCursor(layer, [queryfield, 'SHAPE@']) as cursor:
            for row in cursor:
                for area in Processing_Variables['ShapeArea']:
                    if not row[1].disjoint(area):
                        Processing_Variables['whacodes'].append(row[0])

    concat_attribute = ''
    for zone in Processing_Variables['whacodes']:
        if zone in query:
            Processing_Variables['Applicable'][label] = 'Y'
            if str(zone) not in lst_attributes:
                lst_attributes.append(str(zone))
        else:
            pass

    for attr in lst_attributes:
        concat_attribute += attr
        if attr != lst_attributes[-1]:
            concat_attribute += ', '

    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]
    return ()


# ----------------------------------------------------------------------------------------------------
#
#
# ----------------------------------------------------------------------------------------------------


def UngulateWinterRange(label, query, queryfield, cursorfieldlist, layer):
    lst_attributes = []
    dict_uwr = defaultdict(list)
    Processing_Variables['ungulatecodes'] = []
    if len(Processing_Variables['ungulatecodes']) < 1:
        # Message("Analyzing: Ungulate Winter Range",4)
        lst_fields = [str(queryfield)] + cursorfieldlist
        with arcpy.da.SearchCursor(layer, lst_fields) as cursor:
            for row in cursor:
                for area in Processing_Variables['ShapeArea']:
                    if not row[lst_fields.index('SHAPE@')].disjoint(area):
                        Processing_Variables['ungulatecodes'].append(row[0])
                        dict_uwr[str(row[0])].append(str(row[2]))

    concat_attribute = ''
    for zone in Processing_Variables['ungulatecodes']:
        if zone in query:
            Processing_Variables['Applicable'][label] = 'Y'
            Processing_Variables['CannedStatement'][label] = \
                Processing_Variables['CannedStatement'][label].replace('##', ','.join(dict_uwr[zone]))
        else:
            pass


    # if len(lst_attributes) > 0:
    #     arcpy.AddMessage(concat_attribute)
    #     arcpy.AddMessage(Processing_Variables['CannedStatement'][label])
    #     Processing_Variables['CannedStatement'][label] = \
    #         Processing_Variables['CannedStatement'][label].replace('##', concat_attribute)
    try:
        Processing_Variables['AttributeInformation'][label].append(concat_attribute)
    except:
        Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    return ()


# ----------------------------------------------------------------------------------------------------
#
# The following requires special processing and cannot be run in the above processess
#
# ----------------------------------------------------------------------------------------------------
def SpecialProcessing(label, layer, cursorfieldlist, BufferDistance, layer_name):
    lst_attributes = []

    # ---------------------------------------------------------------------------------------------------------
    # Landscape Level Biodiversity
    if label in ['Landscape Level Biodiversity: Adjacent to Another Harvested Cutblock (100m Buffer Analysis)',
                 'Landscape Level Biodiversity: Adjacent to Another Planned Cutblock (100m Buffer Analysis)']:
        with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
            for row in cursor:
                if str(row[cursorfieldlist.index("SHAPE@")]) == "None":
                    pass
                else:
                    for area in Processing_Variables['ShapeArea']:
                        if row[0].equals(area):
                            pass
                        else:
                            buff = area.buffer(BufferDistance)
                            # if row[0].overlaps(buff) or row[0].touches(buff) or buff.contains(row[0]):
                            if not row[0].disjoint(buff):
                                attribute = ''
                                Processing_Variables['Applicable'][label] = 'Y'
                                if len(cursorfieldlist) == 2:
                                    attribute += '{}'.format(row[1]).replace('Shaha', 'Skaha')
                                elif len(cursorfieldlist) > 2:
                                    for num in range(1, len(cursorfieldlist)):
                                        if num == 2:
                                            attribute += ' ('
                                        attribute += '{}'.format(row[num]).replace('Shaha', 'Skaha')
                                        if 2 <= num < (len(cursorfieldlist) - 1):
                                            attribute += ', '
                                        elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                            attribute += ')'
                                if attribute not in lst_attributes:
                                    lst_attributes.append(attribute)

        concat_attribute = ''
        if len(lst_attributes) > 0:
            delim = ', '
            # if len(cursorfieldlist) == 2:
            #     delim = ', '
            # elif len(cursorfieldlist) > 2:
            #     delim = '\n'

            for attr in lst_attributes:
                concat_attribute += attr
                if attr != lst_attributes[-1]:
                    concat_attribute += delim

        try:
            Processing_Variables['AttributeInformation'][label].append(concat_attribute)
        except:
            Processing_Variables['AttributeInformation'][label] = [concat_attribute]



    # ---------------------------------------------------------------------------------------------------------
    # MaxCutblockSize
    elif label == "Landscape Level Biodiversity: Max Cutblock Size":

        if Processing_Variables['RCTYPE'] == "Block":
            arcpy.AddMessage("yooooooooo")
            totalarea = 0
            # Search for NAR area, if the SU is PROD status.
            with arcpy.da.SearchCursor(layer, ["CUTB_SEQ_NBR", "NAR",
                                               "SUTY_TYPE_ID"]) as cursor:
                for row in cursor:
                    if row[0] == Processing_Variables['UNIQUEID']:
                        if row[2] in ['PROD']:
                            totalarea += float(str(row[1]).replace("None", '0'))

            totalarea = round(totalarea, 1)

            if totalarea > 40:
                Processing_Variables['Applicable'][label] = 'Y'
                Processing_Variables['AttributeInformation'][label] = 'NAR: ' + str(totalarea) + ' ha.'
            else:
                Processing_Variables['Applicable'][label] = 'N'
                Processing_Variables['AttributeInformation'][label] = 'NAR: ' + str(totalarea) + ' ha.'
            arcpy.AddMessage(totalarea)

            # if NAR NUM does not work for the block, then go by Gross Area.
            if totalarea == 0:
                arcpy.AddMessage("hello")
                with arcpy.da.SearchCursor(Processing_Variables['ShapeOfInterest'], ["GROSS_AREA"]) as cursor:
                    for row in cursor:
                        totalarea += float(str(row[0]).replace("None", '0'))
                totalarea = round(totalarea, 1)
                if int(totalarea) > 40:
                    Processing_Variables['Applicable'][label] = 'Y'
                    Processing_Variables['AttributeInformation'][label] = 'Gross Area: ' + str(totalarea) + ' ha.'
                else:
                    Processing_Variables['Applicable'][label] = 'N'
                    Processing_Variables['AttributeInformation'][label] = 'Gross Area: ' + str(totalarea) + ' ha.'

        elif Processing_Variables['RCTYPE'] == "Road":
            Processing_Variables['Applicable'][label] = "I dont know what to do with roads yet."

    elif label in ['Private Land', 'Parks', 'Rare Ecosystems']:
        canned_statement = str(Processing_Variables['CannedStatement'][label]).replace('None', '').split('//')
        if len(canned_statement) > 1:
            overlap_statement = canned_statement[0]
            buff_statement = canned_statement[1]
        else:
            overlap_statement = canned_statement[0]
            buff_statement = canned_statement[0]

        lst_attributes = []
        lst_shapes = []
        with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
            for row in cursor:
                if str(row[0]) == "None":
                    pass
                else:
                    for area in Processing_Variables['ShapeArea']:
                        if not row[0].disjoint(area):
                            attribute = ''
                            Processing_Variables['Applicable'][label] = 'Y'
                            if len(cursorfieldlist) == 2:
                                attribute += '{}'.format(row[1])
                            elif len(cursorfieldlist) > 2:
                                for num in range(1, len(cursorfieldlist)):
                                    if num == 2:
                                        attribute += ' ('
                                    attribute += '{}'.format(row[num])
                                    if 2 <= num < (len(cursorfieldlist) - 1):
                                        attribute += ', '
                                    elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                        attribute += ')'
                            if attribute not in lst_attributes:
                                lst_attributes.append(attribute)
                            lst_shapes.append(row[0])

        concat_attribute = ''
        if len(lst_attributes) > 0:
            delim = ', '
            # if len(cursorfieldlist) == 2:
            #     delim = ', '
            # elif len(cursorfieldlist) > 2:
            #     delim = '\n'

            for attr in lst_attributes:
                concat_attribute += attr
                if attr != lst_attributes[-1]:
                    concat_attribute += delim

        if Processing_Variables['Applicable'][label] == 'Y':
            Processing_Variables['CannedStatement'][label] = overlap_statement
        elif Processing_Variables['Applicable'][label] == 'X -' + str(BufferDistance):
            Processing_Variables['CannedStatement'][label] = buff_statement
            canned_statement = str(Processing_Variables['CannedStatement'][label]).replace('None', '')
            if concat_attribute != '':
                concat_attribute = canned_statement + '\n' + concat_attribute
            else:
                concat_attribute = canned_statement

        try:
            Processing_Variables['AttributeInformation'][label].append(concat_attribute)
        except:
            Processing_Variables['AttributeInformation'][label] = [concat_attribute]

        check_range(layer=layer, label=label, cursorfieldlist=cursorfieldlist, Buffer_Meters=BufferDistance,
                    layer_name=layer_name, lst_shapes=lst_shapes)

    elif label in ['Walk-in Lakes']:
        canned_statement = str(Processing_Variables['CannedStatement'][label]).replace('None', '').split('//')
        if len(canned_statement) > 1:
            overlap_statement = canned_statement[0]
            buff_statement = canned_statement[1]
        else:
            overlap_statement = canned_statement[0]
            buff_statement = canned_statement[0]

        lst_attributes = []
        lst_shapes = []
        with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
            for row in cursor:
                if str(row[0]) == "None":
                    pass
                else:
                    for area in Processing_Variables['ShapeArea']:
                        buff = area.buffer(BufferDistance)
                        if not row[0].disjoint(buff):
                            attribute = ''
                            Processing_Variables['Applicable'][label] = 'Y'
                            if len(cursorfieldlist) == 2:
                                attribute += '{}'.format(row[1])
                            elif len(cursorfieldlist) > 2:
                                for num in range(1, len(cursorfieldlist)):
                                    if num == 2:
                                        attribute += ' ('
                                    attribute += '{}'.format(row[num])
                                    if 2 <= num < (len(cursorfieldlist) - 1):
                                        attribute += ', '
                                    elif num >= 2 and num == (len(cursorfieldlist) - 1):
                                        attribute += ')'
                            if attribute not in lst_attributes:
                                lst_attributes.append(attribute)
                            lst_shapes.append(row[0])

        # if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
        #     with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
        #         for row in cursor:
        #             attribute = ''
        #             for area in Processing_Variables['ShapeArea']:
        #                 buff = area.buffer(BufferDistance)
        #                 if not row[0].disjoint(buff):
        #                     Processing_Variables['Applicable'][label] = 'X -' + str(BufferDistance)
        #                     if len(cursorfieldlist) > 1 and label in ['Parks']:
        #                         if len(cursorfieldlist) == 2:
        #                             attribute += '{}'.format(row[1])
        #                         elif len(cursorfieldlist) > 2:
        #                             for num in range(1, len(cursorfieldlist)):
        #                                 if num == 2:
        #                                     attribute += ' ('
        #                                 attribute += '{}'.format(row[num])
        #                                 if 2 <= num < (len(cursorfieldlist) - 1):
        #                                     attribute += ', '
        #                                 elif num >= 2 and num == (len(cursorfieldlist) - 1):
        #                                     attribute += ')'
        #                         if attribute not in lst_attributes:
        #                             lst_attributes.append(attribute)

        concat_attribute = ''
        if len(lst_attributes) > 0:
            delim = ', '
            # if len(cursorfieldlist) == 2:
            #     delim = ', '
            # elif len(cursorfieldlist) > 2:
            #     delim = '\n'

            for attr in lst_attributes:
                concat_attribute += attr
                if attr != lst_attributes[-1]:
                    concat_attribute += delim

        if Processing_Variables['Applicable'][label] == 'Y':
            Processing_Variables['CannedStatement'][label] = overlap_statement
        elif Processing_Variables['Applicable'][label] == 'X -' + str(BufferDistance):
            Processing_Variables['CannedStatement'][label] = buff_statement
            canned_statement = str(Processing_Variables['CannedStatement'][label]).replace('None', '')
            if concat_attribute != '':
                concat_attribute = canned_statement + '\n' + concat_attribute
            else:
                concat_attribute = canned_statement

        try:
            Processing_Variables['AttributeInformation'][label].append(concat_attribute)
        except:
            Processing_Variables['AttributeInformation'][label] = [concat_attribute]

    # ---------------------------------------------------------------------------------------------------------
    # Consultative Areas
    elif label == 'Consultative Areas':
        with arcpy.da.SearchCursor(layer, ['SHAPE@', "BOUNDARY_NAME", "CONTACT_ORG"]) as cursor:
            for row in cursor:
                if str(row[0]) == "None":
                    pass
                else:
                    for area in Processing_Variables['ShapeArea']:
                        buff = area.buffer(BufferDistance)
                        if not row[0].disjoint(buff):
                            if row[1] == row[2]:
                                name = '    ' + row[1]
                            else:
                                name = '    ' + row[1] + ': ' + row[2]
                            Processing_Variables['Applicable'][label] = 'Y'
                            Processing_Variables['Applicable'][name] = 'Y'
                            Processing_Variables['Sensitive_Information'].append(name)

    # ---------------------------------------------------------------------------------------------------------
    # Grizzly Bear
    elif label in ['Grizzly Bear Habitat', 'Grizzly Bear Habitat RMZ']:
        def grizhabitat():
            with arcpy.da.SearchCursor('LRMP Grizzly Bear RMZ', ['SHAPE@']) as cursor:
                for row in cursor:
                    for area in Processing_Variables['ShapeArea']:
                        if row[0].overlaps(area) or row[0].contains(area) or area.contains(row[0]):
                            Processing_Variables['Applicable']['Grizzly Bear Habitat RMZ'] = 'Y'

        def grizsuitability():
            env.workspace = 'in_memory'
            arcpy.Intersect_analysis(["LRMP Grizzly Bear Suitability", Processing_Variables['ShapeOfInterest']],
                                     'in_memory\\GrizzlySuitIntersect')
            value = {"High": 1, "High-Mod": 2, "Moderate": 3, "Low": 4, "Very Low": 5, "Nil": 6, "Unrated": 7}
            highestrating = 90
            with arcpy.da.SearchCursor('in_memory\\GrizzlySuitIntersect', ['SUIT']) as cursor:
                for row in cursor:
                    try:
                        if value[row[0]] < highestrating:
                            highestrating = value[row[0]]
                    except:
                        pass
            try:
                numtocode = {1: "H", 2: "H-M", 3: 'M', 4: 'L', 5: 'VL', 6: 'Nil', 7: 'N/R'}
                Processing_Variables['Applicable']['Grizzly Bear Habitat'] = 'Y'
                Processing_Variables['AttributeInformation']['Grizzly Bear Habitat'] = "Highest Ranking Value: " + \
                                                                                       numtocode[highestrating]
            except:
                Processing_Variables['Applicable']['Grizzly Bear Habitat'] = 'N'
                return Processing_Variables['Applicable'], Processing_Variables['AttributeInformation']

        # run grizzly bear
        grizhabitat()
        # If griz habitat is applicable run the suitability.
        if Processing_Variables['Applicable']['Grizzly Bear Habitat RMZ'] == 'Y':
            grizsuitability()
        else:
            Processing_Variables['Applicable']['Grizzly Bear Habitat'] = 'N'

    elif label == 'Invasive Plants':
        mapcodes = {'AR': 'African rue / harmal', 'BH': 'Black henbane', 'BW': 'Blueweed',
                    'RI': 'Bog bulrush / ricefield bulrush', 'BO': 'Bohemian knotweed', 'ED': 'Brazilian waterweed',
                    'BT': 'Bull thistle ', 'BU': 'Burdock species ', 'AM': 'Camel thorn', 'CT': 'Canada thistle',
                    'CA': 'Caraway', 'CY': 'Chicory', 'CE': 'Clary sage', 'AO': 'Common bugloss ',
                    'CC': 'Common crupina', 'CX': 'Common hawkweed ', 'RC': 'Common reed', 'TC': 'Common tansy',
                    'CL': 'Cutleaf blackberry ', 'DT': 'Dalmatian toadflax ', 'DC': 'Dense-flowered cordgrass ',
                    'DK': 'Diffuse knapweed', 'DW': 'Dyers woad', 'ES': 'Eggleaf spurge', 'EC': 'English cordgrass',
                    'EH': 'European hawkweed', 'BF': 'False brome', 'FS': 'Field scabious', 'FR': 'Flowering rush',
                    'GL': 'Garden yellow loosestrife', 'AP': 'Garlic mustard', 'GH': 'Giant hogweed',
                    'GK': 'Giant knotweed', 'SW': 'Giant mannagrass / reed sweetgrass', 'AD': 'Giant reed / giant cane',
                    'RG': 'Goats rue / french lilac', 'GO': 'Gorse', 'HR': 'Hairy cats-ear ', 'HS': 'Hawkweed species',
                    'HI': 'Himalayan blackberry', 'PO': 'Himalayan knotweed ', 'HA': 'Hoary alyssum',
                    'HC': 'Hoary cress', 'HY': 'Hydrilla', 'IS': 'Iberian starthistle',
                    'IT': 'Italian plumeless thistle', 'JK': 'Japanese knotweed', 'GJ': 'Johnsongrass ',
                    'JG': 'Jointed goatgrass', 'KH': 'King devil hawkweed', 'KS': 'Knapweed species', 'KU': 'Kudzu',
                    'LS': 'Leafy spurge ', 'MX': 'Maltese star thistle', 'MT': 'Marsh plume thistle/Marsh thistle ',
                    'MC': 'Meadow clary ', 'MH': 'Meadow hawkweed ', 'MK': 'Meadow knapweed ',
                    'MS': 'Mediterranean sage ', 'TM': 'Medusahead', 'MU': 'Mullein', 'NA': 'North africa grass ',
                    'OH': 'Orange hawkweed ', 'OD': 'Oxeye daisy', 'PF': 'Parrot feather', 'PP': 'Perennial pepperweed',
                    'PS': 'Perennial sow thistle ', 'PA': 'Polar hawkweed', 'IM': 'Policemans helmet / him. balsam',
                    'PV': 'Puncturevine ', 'PN': 'Purple nutsedge ', 'PU': 'Purple starthistle ',
                    'QH': 'Queen devil hawkweed', 'BR': 'Red bartsia', 'RS': 'Rush skeletonweed',
                    'AH': 'Saltlover / halogeton ', 'SN': 'Salt-meadow cord grass', 'SA': 'Saltwater cord grass',
                    'SH': 'Scentless chamomile', 'SB': 'Scotch broom ', 'ST': 'Scotch thistle', 'SG': 'Shiny geranium',
                    'NS': 'Silverleaf nightshade ', 'FT': 'Slender meadow foxtail', 'SM': 'Smooth hawkweed ',
                    'SO': 'Sowthistle species ', 'SX': 'Spotted hawkweed', 'SK': 'Spotted knapweed',
                    'MV': 'Spring millet grass', 'TP': 'Spurge flax', 'CV': 'Squarrose knapweed ',
                    'SC': 'Sulphur cinquefoil ', 'SY': 'Syrian bean-caper', 'TH': 'Tall hawkweed',
                    'TR': 'Tansy ragwort', 'TS': 'Teasel ', 'TX': 'Texas blueweed', 'VL': 'Velvet leaf',
                    'WA': 'Wall hawkweed', 'TN': 'Water chestnut', 'WH': 'Water hyacinth', 'AQ': 'Water soldier',
                    'WG': 'Western goats-beard', 'WP': 'Whiplash hawkweed', 'WI': 'Wild chervil ',
                    'WM': 'Wild mustard ', 'WT': 'Winged / slender-flowered thistle', 'YA': 'Yellow archangel',
                    'YD': 'Yellow devil hawkweed ', 'YH': 'Yellow hawkweed ', 'YI': 'Yellow iris <5m2',
                    'YN': 'Yellow nutsedge ', 'YS': 'Yellow starthistle ', 'AB': 'American beachgrass',
                    'YC': 'Amphibious yellow cress', 'HB': 'Annual hawksbeard', 'AS': 'Annual sow thistle ',
                    'BY': 'Babys breath', 'BB': 'Bachelors button', 'BA': 'Barnyard grass', 'KB': 'Bighead knapweed',
                    'BP': 'Bigleaf / Large periwinkle', 'BL': 'Black knapweed', 'RB': 'Black locust ',
                    'BC': 'Bladder campion ', 'RA': 'Bristly locust / rose acacia', 'BK': 'Brown knapweed',
                    'CB': 'Bur chervil', 'BD': 'Butterfly bush', 'CG': 'Carpet burweed',
                    'DB': 'Cheatgrass / downy brome ', 'LC': 'Cherry laurel', 'CH': 'Chilean tarweed ',
                    'CF': 'Coltsfoot ', 'CO': 'Common comfrey', 'FC': 'Common frogbit', 'CP': 'Common periwinkle',
                    'CR': 'Creeping buttercup ', 'CU': 'Cudweed', 'CD': 'Curled dock', 'UP': 'Curly leaf pondweed',
                    'CS': 'Cypress spurge', 'DR': 'Dames rocket', 'SL': 'Daphne / spurge laurel', 'DI': 'Didymo ',
                    'DO': 'Dodder ', 'DE': 'Dwarf eelgrass', 'HO': 'English holly', 'EI': 'English ivy',
                    'EW': 'Eurasian watermilfoil ', 'EB': 'European beachgrass', 'EL': 'European lake sedge',
                    'MQ': 'European water clover ', 'WE': 'European waterlily ', 'EY': 'Eyebright ', 'FW': 'Fanwort',
                    'FM': 'Feathered mosquito-fern', 'FB': 'Field bindweed', 'FP': 'Flat pea / flat peavine',
                    'FL': 'Fragrant water lily', 'GM': 'French broom ', 'MA': 'Giant chickweed ',
                    'SV': 'Giant salvinia', 'GW': 'Goutweed / bishops weed ', 'GC': 'Greater celandine',
                    'GN': 'Greater knapweed', 'GF': 'Green foxtail / green bristlegrass', 'GS': 'Groundsel ',
                    'BI': 'Hedge false bindweed', 'HD': 'Hedgehog dogtail', 'GR': 'Herb robert', 'HT': 'Hounds-tongue',
                    'JW': 'Japanese wireweed', 'KO': 'Kochia ', 'LT': 'Ladys-thumb ',
                    'LL': 'Large yellow / spotted loosestrife', 'RF': 'Lesser celandine / fig buttercup',
                    'LO': 'Longspine sandbur', 'OW': 'Major oxygen weed', 'MB': 'Meadow buttercup',
                    'MG': 'Meadow goats-beard ', 'MI': 'Milk thistle ', 'MO': 'Mountain bluet',
                    'ME': 'Mouse ear hawkweed ', 'NC': 'Night-flowering catchfly ', 'NI': 'Nightshade',
                    'NT': 'Nodding thistle ', 'OM': 'Old mans beard / travellers joy ', 'PT': 'Plumeless thistle',
                    'PH': 'Poison hemlock', 'LP': 'Portugese laurel', 'PR': 'Portuguese broom',
                    'PC': 'Prickly comfrey ', 'PD': 'Purple deadnettle', 'PL': 'Purple loosestrife ',
                    'QA': 'Queen annes lace / wild carrot', 'RP': 'Redroot amaranth / rough pigweed',
                    'RK': 'Russian knapweed', 'RO': 'Russian olive', 'RT': 'Russian thistle ',
                    'TA': 'Saltcedar / tamarisk', 'SS': 'Sheep sorrel ', 'SP': 'Shepherds-purse',
                    'CN': 'Short-fringed knapweed', 'SE': 'Siberian elm ', 'HG': 'Smooth cats ear',
                    'BS': 'Spanish bluebells', 'SI': 'Spanish broom', 'SJ': 'St. Johns wort/Goatweed',
                    'SF': 'Sweet fennel ', 'TB': 'Tartary buckwheat', 'AA': 'Tree of heaven',
                    'LM': 'Variable leaf milfoil ', 'WL': 'Wand loosestrife', 'LW': 'Water lettuce', 'NO': 'Watercress',
                    'WC': 'White cockle ', 'WB': 'Wild buckwheat', 'WF': 'Wild four oclock', 'WO': 'Wild oats ',
                    'PW': 'Wild parsnip ', 'WS': 'Wood sage ', 'WW': 'Wormwood', 'YF': 'Yellow floating heart ',
                    'YT': 'Yellow/common toadflax'}

        lst_attributes = []
        with arcpy.da.SearchCursor(layer, ['SHAPE@', 'MAP_LABEL']) as cursor:
            for row in cursor:
                if str(row[0]) == "None":
                    pass
                else:
                    for area in Processing_Variables['ShapeArea']:
                        if row[0].overlaps(area) or row[0].contains(area) or area.contains(row[0]):
                            Processing_Variables['Applicable'][label] = 'Y'
                            for item in row[1].split(' '):
                                commonname = mapcodes[item]
                                if commonname not in lst_attributes:
                                    lst_attributes.append(commonname)

        concat_attribute = ''
        for attr in lst_attributes:
            concat_attribute += attr
            if attr != lst_attributes[-1]:
                concat_attribute += ', '

        try:
            Processing_Variables['AttributeInformation'][label].append(concat_attribute)
        except:
            Processing_Variables['AttributeInformation'][label] = [concat_attribute]

        if Processing_Variables['Applicable'][label][:1] not in ['X', 'Y']:
            with arcpy.da.SearchCursor(layer, cursorfieldlist) as cursor:
                for row in cursor:
                    for area in Processing_Variables['ShapeArea']:
                        buff = area.buffer(BufferDistance)
                        if not row[0].disjoint(buff):
                            Processing_Variables['Applicable'][label] = 'X -' + str(BufferDistance)
                            break

    else:
        pass

    return


# ----------------------------------------------------------------------------------------------------
#
# The following requires special processing and cannot be run in the above processess
#
# ----------------------------------------------------------------------------------------------------
def SpecialProcessingNoLayer(label):
    # ---------------------------------------------------------------------------------------------------------
    # Hydrological
    if label == 'Hydrological':
        for item in ["Fisheries Sensitive Watershed", 'Community Watersheds']:
            try:
                if Processing_Variables['Applicable'][item] == 'Y':
                    Processing_Variables['Applicable'][label] = 'Y'
                    return ()
            except:
                pass

    return


# ========================================================================================================
#  Initial Routines
# ========================================================================================================
def OrganizeRouteCardItems():
    Processing_Variables['PROCESS_STATUS'] = 0
    Processing_Variables['PROCESS_INFO'] = "PreppingMXD"

    Message('    - Running Analysis', 2)

    # Sort Processing Table in the order that items have been assigned.
    Processing_Variables["Applicable"] = OrderedDict()
    arcpy.Sort_management(Processing_Variables['ProcessingLookupTable'], 'in_memory\\sorttable',
                          [[Processing_Variables['LegalLocation'], "ASCENDING"]])

    Processing_Variables['AttributeInformation'].clear()
    Processing_Variables['CannedStatement'].clear()
    # del Processing_Variables['ExistingBEC'][:]

    itemlist = []
    cursorfieldlist = ["Item", "Processing", "Layer_List", "Query_Values", "Query_Field", "CannedStatement",
                       "BufferDistance", 'Sensitive_Info', 'Definition_Query']
    cursorfieldlist.append(Processing_Variables['LegalLocation'])
    for row in arcpy.da.SearchCursor('in_memory\\sorttable', cursorfieldlist):
        if row[cursorfieldlist.index(Processing_Variables['LegalLocation'])] > 0:
            # Create the list of all items taht should be used based on the location.
            item = row[cursorfieldlist.index("Item")]
            Processing = row[cursorfieldlist.index("Processing")]
            Layer_List = str(row[cursorfieldlist.index("Layer_List")]).replace("None", '').split(';')
            Query_Values = str(row[cursorfieldlist.index("Query_Values")]).replace("None", '').split(',')
            Query_Field = row[cursorfieldlist.index("Query_Field")]
            Sensitive_Label = row[cursorfieldlist.index("Sensitive_Info")]
            Definition_Query = row[cursorfieldlist.index('Definition_Query')]
            try:
                Buffer_Meters = float(row[cursorfieldlist.index("BufferDistance")])
            except:
                # Buffer_Meters = float(Processing_Variables['Variables']['SearchBufferDistance'])
                Buffer_Meters = 0

            for lyr in Layer_List:
                if lyr in brknList:
                    # Processing_Variables['Applicable'][item] = 'ACCESS DENIED'
                    # Processing_Variables['CannedStatement'][item] = 'User does not have permissions to access this layer'
                    Processing = 'accessdenied'

            # LegendGroup = row[cursorfieldlist.index("LegendGroup")]
            Processing_Variables['CannedStatement'][item] = row[cursorfieldlist.index("CannedStatement")]
            # LegendElement = row[cursorfieldlist.index("LegendElement")]
            itemlist.append([item, Processing, Layer_List, Query_Values, Query_Field, Buffer_Meters, Definition_Query])
            Processing_Variables['layerdict'][item] = [Layer_List]
            if Sensitive_Label == 'Y':
                Processing_Variables['Sensitive_Information'].append(item)

    for i in itemlist:
        label = i[0]
        processingunit = i[1]

        if processingunit == 'title':
            output = ''
        elif processingunit == 'nonspatial':
            output = 'ZZZ'
        elif processingunit == 'accessdenied':
            output = 'ACCESS DENIED'
            Processing_Variables['CannedStatement'][label] = 'User does not have permissions to access this layer or this layer no longer exists'
        else:
            output = 'N'
        Processing_Variables['Applicable'][label] = output
        if label == access_label:
            Processing_Variables['Applicable'][label] = 'N'
            Processing_Variables['AttributeInformation'][label] = ''

    ungulatecount = 0
    whacount = 0
    bandcount = 0


    # Start the Processing Options
    for i in itemlist:
        label = i[0]
        processingunit = i[1]
        buffdist = i[5]
        def_query = i[6]
        if processingunit in ('title', 'nonspatial'):
            pass
        else:
            layerslist = i[2]
            for layer in layerslist:
                cursorlist = ['SHAPE@']
                # Message( "Analyzing: "+ label,2)
                # if there is information that is wanted from the layer, extract fields of interest.
                if ':' in layer:
                    # fields split must occur before the layer split because we are reassigning layer
                    fields = layer.split(':')[1].split(',')
                    layer = layer.split(':')[0]

                    if len(fields) > 0:
                        for f in fields:
                            if f == '':
                                pass
                            else:
                                f_split = f.split('=')
                                if len(f_split) > 1:
                                    dict_labels[f_split[1]] = f_split[0]
                                    cursorlist.append(f_split[1])
                                else:
                                    dict_labels[f] = f
                                    cursorlist.append(f)
                # Skip values that have already been determined
                if Processing_Variables['Applicable'][label] in ['BA']:
                    pass

                # Contained if outside of the regular functions due to the fact that we need to bypass the "NoOverlap" Section.
                if processingunit == "contained":
                    contained(layer, label, cursorlist)
                    arcpy.AddMessage(label)
                    arcpy.AddMessage(processingunit)

                if layer in Processing_Variables['NoOverlap']:
##                    arcpy.AddMessage(label)
##                    arcpy.AddMessage(processingunit)
                    pass

                else:
                    try:
                        # send to  analysis where the layer is predefined and hardcoded into the script
                        if processingunit == "special_processing_no_layer":
                            SpecialProcessingNoLayer(label)
                        if processingunit == "default_to_yes":
                            Processing_Variables['Applicable'][label] = 'Y'

                        # define layer for analysis
                        try:
                            selectname = Processing_Variables['SelectionLayers'][layer]
                            selectionlayer = 'in_memory\\' + selectname
                        except:
                            selectionlayer = layer

                        # if layer in ['Old Forest In BEC < 10% Old Forest', 'Ancient Forest Probability',
                        #              'Old Forest Where Provincial Site Productivity Site Index >20m',
                        #              'Old Forest Where VRI Site Index >20m']:
                        #     if processingunit != 'accessdenied':
                        #         selectionlayer = 'in_memory\\oldforest'
                        #         arcpy.MultipartToSinglepart_management(in_features=layer, out_feature_class=selectionlayer)

                        if processingunit == "special_processing":
##                            arcpy.AddMessage("Special Processing")
##                            arcpy.AddMessage(label)

                            SpecialProcessing(label, selectionlayer, cursorlist, buffdist, layer)

                        if processingunit == 'contains_overlap_layers':
                            containsoverlap(selectionlayer, label, cursorlist, buffdist, def_query, layer, len(layerslist))

                        if processingunit == 'SOMC':
                            SOMC(selectionlayer, label, cursorlist, buffdist, def_query, layer, len(layerslist))

                        if processingunit == 'highestvalue':
                            containsoverlap(selectionlayer, label, cursorlist)

                        # if processingunit == 'polygon_coverages':
                        #     polygoncoverages(selectionlayer, label)

                        if processingunit == 'point_line_layers':
                            pointline(selectionlayer, label, cursorlist, buffdist, layer, len(layerslist))

                        if processingunit == 'overlap_touching_layers':
                            overlaptouching(selectionlayer, label, cursorlist, buffdist)

                        if processingunit == 'BEC':
                            query = i[3]
                            field = i[4]
                            BECDefinedAreas(label, query, field, selectionlayer)

                        if processingunit == 'ungulate_winter_range':
                            if ungulatecount > 0 and len(Processing_Variables['ungulatecodes']) < 1:
                                Processing_Variables['Applicable'][label] = 'N'
                            else:
                                ungulatecount += 1
                                query = i[3]
                                field = i[4]
                                UngulateWinterRange(label, query, field, cursorlist, selectionlayer)

                        if processingunit == 'wildlife_habitat_areas':
                            if whacount > 0 and len(Processing_Variables['whacodes']) < 1:
                                Processing_Variables['Applicable'][label] = 'N'
                            else:
                                whacount += 1
                                query = i[3]
                                field = i[4]
                                WildlifeHabitatAreas(label, query, field, selectionlayer)
                    except:

                        e = sys.exc_info()[1]
                        arcpy.AddError('Processing Failed: {}: {}'.format(label, layer))
                        arcpy.AddError(e.args[0])
                        sys.exit()
        if (label in lst_constraints) and (access_label in Processing_Variables['Applicable']):
            if Processing_Variables['Applicable'][label] == 'Y':
                Processing_Variables['Applicable'][access_label] = 'Y'
                Processing_Variables['AttributeInformation'][access_label] += '{}, '.format(label)
    becs = str(Processing_Variables['ExistingBEC']).replace('[', '').replace(']', '').replace('u\'', '')\
        .replace('\'', '').replace(' ', '').split(',')

    if access_label in Processing_Variables['Applicable']:
        for bec in becs:
            for const in lst_constraints:
                if const.startswith('BEC'):
                    bec_const = const.split(':')[1]
                    if bec.startswith(bec_const):
                        Processing_Variables['Applicable'][access_label] = 'Y'
                        Processing_Variables['AttributeInformation'][access_label] += 'BEC {}'.format(bec)
                        break

    access_attribute = str(Processing_Variables['AttributeInformation'][access_label])
    if access_attribute.endswith(', '):
        Processing_Variables['AttributeInformation'][access_label] = access_attribute[:-2]

# =========================================================================================================
#  Output Routines
# =========================================================================================================
def WriteToExcel():
    Message('   Create Excel Output', 2)

    wb = xlwt.Workbook()
    for sheet_type in ['', '_Sensitive_Information']:

        if sheet_type == '':
            ws = wb.add_sheet('Route Card Tool Assessment')
        else:
            ws = wb.add_sheet('Consultative Areas')
        ws.set_portrait(False)

        # Set Excel Styles
        xlwt.add_palette_colour("custom_colour", 0x21)
        xlwt.add_palette_colour("darkergreen", 0x22)
        xlwt.add_palette_colour("lightred", 0x23)
        wb.set_colour_RGB(0x21, 216, 228, 188)
        wb.set_colour_RGB(0x22, 196, 215, 155)
        wb.set_colour_RGB(0x23, 248, 203, 173)

        BigTitle = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold on; pattern: pattern solid, fore_colour darkergreen; align: wrap on, vert center;')
        BlackBoldStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold on; pattern: pattern solid, fore_colour custom_colour; align: wrap on, vert center;')
        RedStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index red; align: wrap on, vert center; borders: left thin, right thin, top thin, bottom thin, bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        RedBoldStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index red, bold on; pattern: pattern solid, fore_colour lightred; align: wrap on, vert center; borders: left thin, right thin, top thin, bottom thin, bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        BlackStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index black; align: wrap on, vert center;borders: left thin, right thin, top thin, bottom thin, bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        BlackStyleNoBorder = xlwt.easyxf('font: name Times New Roman, color-index black; align: wrap on, vert center;')
        OrangeStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index orange; align: wrap on; borders: left thin, right thin, top thin, bottom thin, bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        BlueStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index blue; align: wrap on, vert center; borders: left thin, right thin, top thin, bottom thin,bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        GreyStyle = xlwt.easyxf(
            'font: name Times New Roman, color-index gray50; align: wrap on, vert center;borders: left thin, right thin, top thin, bottom thin,bottom_colour gray25, left_colour gray25, right_colour gray25, top_colour gray25;')
        GreyStyleBottomOnly = xlwt.easyxf(
            'font: name Times New Roman,height 260, color-index gray50; align: wrap on, vert center; borders: bottom thin, bottom_colour gray25;')
        OrangeStyle2 = xlwt.easyfont('name Times New Roman, color_index orange')

        # Set the column Widths

        ws.col(0).width = 256 * 5
        ws.col(1).width = 256 * 37
        ws.col(2).width = 256 * 15
        ws.col(3).width = 256 * 20
        ws.col(4).width = 256 * 60
        ws.col(5).width = 256 * 8

        # --------------------------------------
        # START: Headers and Titles Below
        # --------------------------------------
        outputname = ''
        for item in Processing_Variables['TitleFields'].split(','):
            outputname += str(Processing_Variables['Title'][item]) + "_"

        ws.write_merge(0, 0, 1, 3, 'FRPA Planning Route Card - ' + str(Processing_Variables['LegalLocation']),
                       GreyStyleBottomOnly)
        ws.write_merge(2, 2, 1, 4, outputname.replace('_', ' '), BigTitle)

        ws.row(5).height_mismatch = True
        ws.row(5).height = 760
        ws.write_merge(5, 5, 1, 4,
                       "OBJECTIVE: Planning Route Card is a due diligence checklist for the Planning Forester to identify legal and non-legal commitments associated with proposed block/road development, to track completion of assessments/analyses, to identify roles and responsibilities, and provide a communication tools between the Planning Foresters, Practices Foresters, and layout contractors.",
                       BlackStyleNoBorder)

        # --------------------------------------
        # END: Headers and Titles Below
        # --------------------------------------
        # write (Row, Column, information, style)
        rownum = 7
        applicablecolumn = 2
        commentcolumn = 4
        otherColumn = 3
        titlecolumn = 1

        # Add BEC Values to Output
        rownum = rownum + 1
        becs = str(Processing_Variables['ExistingBEC']).replace('[', '').replace(']', '').replace('u\'', '').replace(
            '\'', '')
        ws.write(rownum, titlecolumn, 'Biogeoclimatic Zone', GreyStyle)
        ws.write_merge(rownum, rownum, applicablecolumn, commentcolumn, becs, GreyStyle)

        # Message(Processing_Variables['Applicable'],2)
        if sheet_type == '_Sensitive_Information':
            applicable_keys = [v for v in Processing_Variables['Applicable'].keys() if
                               v in Processing_Variables['Sensitive_Information']]
        elif sheet_type == '':
            applicable_keys = [v for v in Processing_Variables['Applicable'].keys() if
                               v not in Processing_Variables['Sensitive_Information']]

        for value in Processing_Variables['Applicable']:
            if value in applicable_keys:
                arcpy.AddMessage(value)
                # comment = ''
                comment_list = []

                if Processing_Variables['Applicable'][value][:1] not in ['', 'Z', 'N', 'X'] and value not in [access_label]:
                    if value in Processing_Variables['CannedStatement'].keys():
                        comment = str(Processing_Variables['CannedStatement'][value]).replace('None', '').replace('\u2019', '')
                        if comment != '':
                            comment_list.append(comment)

                if Processing_Variables['Applicable'][value][:1] not in ['', 'Z', "N", "X"]:
                    if value in Processing_Variables['AttributeInformation'].keys():
                        if "Highest Ranking Value" in Processing_Variables['AttributeInformation'][value]:
                            comment_list.append(Processing_Variables['AttributeInformation'][value])
                        elif value in [access_label, 'Landscape Level Biodiversity: Max Cutblock Size']:
                            comment_list.append(Processing_Variables['AttributeInformation'][value])
                        else:
                            lst_cmts = sorted(list(set(Processing_Variables['AttributeInformation'][value])))
                            lst_cmts.sort(key=natural_keys)
                            for val in lst_cmts:
                                    # comment = comment + chr(10) + string
                                    if val != '':
                                        comment_list.append(val)
                if Processing_Variables['Applicable'][value][:1] in ['X']:
                    if value in Processing_Variables['AttributeInformation']: #and value in ['Walk-in Lakes',
                                                                                           # 'Private Land', 'Parks',
                                                                                           # 'Rare Ecosystems']:
                        lst_cmts = sorted(list(set(Processing_Variables['AttributeInformation'][value])))
                        lst_cmts.sort(key=natural_keys)
                        for val in lst_cmts:
                        # for val in sorted(list(set(Processing_Variables['AttributeInformation'][value]))):
                            # comment = comment + chr(10) + string
                            if val != '':
                                comment_list.append(val)

                otherassessment = ' '
                if Processing_Variables['Applicable'][value][:3] == "App":
                    style = RedStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value][:3] == "Not":
                    style = BlackStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value] == '':
                    rownum += 1
                    style = BlackBoldStyle
                    ynvalue = "Applicable (Y/N)"
                    otherassessment = "Additional Assessments Needed (Y/N)"
                    comment = "Comments"
                elif Processing_Variables['Applicable'][value] == "Y":
                    Processing_Variables['Applicable'][value] = "Y"
                    if value == access_label:
                        comment = str(Processing_Variables['CannedStatement'][value]).replace('None', '')
                        if comment != '':
                            comment_list.append(comment)
                    style = RedStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value] == "N":
                    Processing_Variables['Applicable'][value] = "N"
                    style = BlackStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                    if value == 'Hydrological' and Processing_Variables['LegalLocation'] == 'Okanagan':
                        comment_list.append('Check most recent Landbase Reporting ECA analysis to evaluate risk')
                    if value == 'Landscape Level Biodiversity: Max Cutblock Size':
                        comment_list.append(Processing_Variables['AttributeInformation'][value])
                elif Processing_Variables['Applicable'][value] == "ZZZ":
                    Processing_Variables['Applicable'][value] = "Non Spatial"
                    style = BlackStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value] == "SOMC":
                    Processing_Variables['Applicable'][value] = "BEC Applicable"
                    style = BlueStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value] == "BA":
                    Processing_Variables['Applicable'][value] = "BEC Applicable"
                    style = BlueStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value][:1] == "X":
                    dist = Processing_Variables['Applicable'][value].split('-')[1]
                    Processing_Variables['Applicable'][value] = "Within " + str(dist).replace('.0', '') + "m of Block"
                    style = OrangeStyle
                    ynvalue = Processing_Variables['Applicable'][value]

                elif Processing_Variables['Applicable'][value] == "NBA":
                    Processing_Variables['Applicable'][value] = "N"
                    comment = str(Processing_Variables['CannedStatement'][value]).replace('None', '')
                    if comment != '':
                        comment_list.append(comment)

                    style = BlackStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                elif Processing_Variables['Applicable'][value] == "ACCESS DENIED":
                    style = RedBoldStyle
                    ynvalue = Processing_Variables['Applicable'][value]
                else:
                    style = GreyStyle
                    ynvalue = Processing_Variables['Applicable'][value]

                rownum += 1

                comment = ''
                for cmt in comment_list:
##                    arcpy.AddMessage(cmt)
                    if cmt != '':
                        comment += cmt
                        if cmt != comment_list[-1]:
                            comment += '\n'

                # write the cells
                ws.write(rownum, titlecolumn, value, style)
                ws.write(rownum, applicablecolumn, ynvalue, style)
                ws.write(rownum, otherColumn, otherassessment, style)
                # ws.write(rownum, commentcolumn, comment, style)

                tpl_cmt = list()
                for cmt in comment_list:
                    if cmt != '':
                        if cmt != comment_list[-1]:
                            cmt += '\n'
                        if cmt.startswith('Within '):
                            tpl_cmt.append((cmt, OrangeStyle2))
                        else:
                            tpl_cmt.append(cmt)
                ws.write_rich_text(rownum, commentcolumn, tpl_cmt, style)
        # Add Disclaimer Statement
        rownum += 2
        ws.row(rownum).height_mismatch = True
        ws.row(rownum).height = 1100
        ws.write_merge(rownum, rownum, 1, 4,
                       'The Planning Route Card (PRC) is a guidance tool listing constraints applicable at time of preparation.  The PRC does not prescribe management practices.  It is incumbent on the Site Plan author and signatory to ensure management practices and decisions are consistent with the intent of the applicable legislation, higher level plans, Forest Stewardship Plan, Statutory Decision Maker direction, best management practices, and general wildlife measures outlined in GAR Orders. The content of the PRC alone cannot be used for justification of, or as a rationale for management decisions contained in any professional documents.',
                       BlackStyle)

        # Add Signature Area
        rownum += 1
        ws.write(rownum, 1, 'Planning Forester Signoff', BlackStyle)
        ws.write_merge(rownum, rownum, 2, 3, 'Prepared By:  ' + Processing_Variables['Owner'], BlackStyle)
        ws.write(rownum, 4, 'Date:  {}'.format(dt.now().strftime('%B %d, %Y')), BlackStyle)

        # Add Map Attachment Block.
        rownum += 2
        ws.write_merge(rownum, rownum, 1, 4,
                       'Attach 1:10,000 letter size map showing block with overlaps and forward to Practices Forester',
                       BlackStyle)

    # Need to change where items will be saved.
    outloc = Processing_Variables['OutputLocation']
    outname = outputname.replace('.0', '').replace('\\', '').replace('/', '')
    status = Processing_Variables['Status']

    outputfile = '{}\\{}{}_RouteCard.xls'.format(outloc, outname, status)
    wb.save(str(outputfile))


# -----------------------------------------------------------------------------------------------------
def Layout():
    Message('   Create Final Map Output', 2)
    mxd = arcpy.mapping.MapDocument("CURRENT")

    Processing_Variables['LayerList']
    # find all layers that have a "hit" so they can be made visible.
    visiblelayers = []

    for l in Processing_Variables['LayerList']:
        if l not in Processing_Variables['NoOverlap']:
            visiblelayers.append(l)

    GrpLyrList = []
    GrpLyrLast = ''
    # Clear selected layers and remove the UBI of Interest layer in memory Turn on any applicable layers.
    for layer in arcpy.mapping.ListLayers(mxd):
        if layer.isGroupLayer:
            GrpLyrLast = layer.name
        if layer.supports("TRANSPARENCY"):
            if layer.transparency > 0:
                layer.transparency = 0
        if layer.name in visiblelayers:
            layer.visible = True
            GrpLyrList.append(GrpLyrLast)
        if layer.name == "SearchBuffer":
            layer.name = str(Processing_Variables['Variables']['SearchBufferDistance']) + ' Meter Search Buffer'
            layer.visible = True
        if layer.name == "ShapeOfInterest":
            layer.name = Processing_Variables['RCTYPE'] + ' of Interest'
            layer.visible = True
        if 'TOC FN' in layer.name or layer.name[:14] == "SensitiveData_":
            layer.visible = False
        else:
            pass

    # Going to activate any group layers that contained layers we turned on above
    for layer in arcpy.mapping.ListLayers(mxd):
        if layer.isGroupLayer:
            if layer.name in (set(GrpLyrList)):
                layer.visible = True

    prevtext = {}
    for lyt in arcpy.mapping.ListLayoutElements(mxd, 'TEXT_ELEMENT'):
        if lyt.name == 'Required':
            pass
        elif lyt.name in Processing_Variables['Title'].keys():
            prevtext[lyt.name] = lyt.text
            lyt.text = lyt.text + " " + str(Processing_Variables['Title'][lyt.name])
        elif lyt.name == 'Owner':
            lyt.text = Processing_Variables['Owner']
        else:
            prevtext[lyt.name] = lyt.text
            lyt.text = " "

    outputname = ''
    for item in Processing_Variables['TitleFields'].split(','):
        outputname += str(Processing_Variables['Title'][item]) + "_"
    arcpy.mapping.ExportToPDF(mxd,
                              Processing_Variables['OutputLocation'] + '\\' + outputname.replace('.0', '').replace('\\',
                                                                                                                   '').replace(
                                  '/', '') + Processing_Variables['Status'] + "_Map.pdf",
                              picture_symbol="VECTORIZE_BITMAP")

    # Reset all title elements and extent!
    for lyt in arcpy.mapping.ListLayoutElements(mxd, 'TEXT_ELEMENT'):
        if lyt.name in prevtext.keys():
            lyt.text = prevtext[lyt.name]

    extlayer = arcpy.mapping.Layer(Processing_Variables['LegalArea'])
    ext = extlayer.getExtent()

    newXMax = ext.XMax
    newXMin = ext.XMin
    newYMax = ext.YMax
    newYMin = ext.YMin

    # Set Processing Extent
    arcpy.env.extent = arcpy.Extent(newXMin, newYMin, newXMax, newYMax)

    # mxd.save()


def atoi(text):
    return int(text) if text.isdigit() else text

def natural_keys(text):
    '''
    alist.sort(key=natural_keys) sorts in human order
    http://nedbatchelder.com/blog/200712/human_sorting.html
    (See Toothy's implementation in the comments)
    '''
    return [ atoi(c) for c in re.split(r'(\d+)', text) ]

# *********************************************************************************************************
#  Main Program
# *********************************************************************************************************

# ---------------------------------------------------------------------------------------------------------
# Run Main for ArcGIS 10.1
#
#   Main Method manages the overall program
# ---------------------------------------------------------------------------------------------------------
def runMain():
    Initialize()
    if RCTYPE == 'Block':
        input_fc = Processing_Variables['Variables']['InputBlockFeatureClass']
    elif RCTYPE == 'Road':
        input_fc = Processing_Variables['Variables']['InputRoadFeatureClass']

    desc = arcpy.Describe(input_fc)
    if str(desc.FIDSet) == "":
        Message("   - Nothing Selected - Stopping Tool", 0)
        sys.exit(100)
    select_fc = Processing_Variables['InputFeatureClass']
    with arcpy.da.SearchCursor(input_fc, 'OID@') as s_cursor:
        for row in s_cursor:
            initialize_variables()
            sql = '{} = {}'.format(arcpy.Describe(input_fc).OIDFieldName, row[0])
            arcpy.Select_analysis(in_features=input_fc, out_feature_class=select_fc, where_clause=sql)
            ResetMXD()
            if RCTYPE == 'Block':
                SelectBlocks()
            elif RCTYPE == 'Road':
                SetupRoads()
            PrepMXD()
            DetermineRequirements()
            SelectLayers()
            OrganizeRouteCardItems()
            WriteToExcel()
            Layout()

    return Processing_Variables['PROCESS_STATUS']


# *********************************************************************************************************

#  Execution
# *********************************************************************************************************

runMain()

# ---------------------------------------------------------------------------------------------------------
#  Clean-up
# ---------------------------------------------------------------------------------------------------------


END_TIME = time.ctime(time.time())
END_TIME_SEC = time.time()

END_TIME_SQL = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

EXECUTION_TIME = END_TIME_SEC - START_TIME_SEC

if EXECUTION_TIME < 120:

    EXECUTION_TIME_STRING = str(int(EXECUTION_TIME)) + ' seconds.'
elif EXECUTION_TIME < 3600:
    EXECUTION_TIME_STRING = str(round(EXECUTION_TIME / 60, 1)) + ' minutes.'
else:
    EXECUTION_TIME_STRING = str(round(EXECUTION_TIME / 3600, 2)) + ' hours.'

# *********************************************************************************************************
#  End of Program
# *********************************************************************************************************
