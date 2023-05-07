#-------------------------------------------------------------------------------
# Name:        Overlap_FN
# Purpose:	Allows the user to quickly select a feature in Arcmap and see 		which FN boundaries it overlaps
#
# Author:      dotto
#
# Created:     15/06/2021
# Copyright:   (c) dotto 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import arcpy,sys,logging,os
from argparse import ArgumentParser

arcpy.Delete_management("in_memory")
arcpy.env.overwriteOutput = True
#   Set the workspace to be in memory so no extra data is created
arcpy.env.workspace = 'in_memory'

sys.path.insert(1, r'W:\FOR\RSI\TOC\Projects\ESRI_Scripts\Python_Repository')
from environment import Environment

sde_folder = 'Database Connections'
input_lyr = arcpy.GetParameterAsText(0)
b_un = arcpy.GetParameterAsText(1)
b_pw = arcpy.GetParameterAsText(2)

bcgw_db = Environment.create_bcgw_connection(location=sde_folder, bcgw_user_name=b_un,
                                                      bcgw_password=b_pw)


FN_Bndry = os.path.join(bcgw_db, 'WHSE_ADMIN_BOUNDARIES.PIP_CONSULTATION_AREAS_SP')

arcpy.MakeFeatureLayer_management(input_lyr,r"in_memory\input_feature")
arcpy.MakeFeatureLayer_management(FN_Bndry,"in_memory\FN_feature")

arcpy.SelectLayerByLocation_management(r"in_memory\FN_feature", 'INTERSECT', r"in_memory\input_feature", selection_type="NEW_SELECTION")

result = arcpy.GetCount_management(r"in_memory\input_feature")
count = str(result.getOutput(0))
arcpy.AddMessage("-------------------------------------------------------------------------------------------------------------------------------------------------")

arcpy.AddMessage("You have {} feature(s) selected from {}. Below are the FN overlaps for these feature(s)".format(count,input_lyr))

arcpy.AddMessage("-------------------------------------------------------------------------------------------------------------------------------------------------")
arcpy.AddMessage("{:50}{:5}{}".format("Contact Organization","|","Consultation Area Name"))
arcpy.AddMessage("-------------------------------------------------------------------------------------------------------------------------------------------------")


for row in arcpy.da.SearchCursor("in_memory\FN_feature",['CONTACT_ORGANIZATION_NAME', 'CNSLTN_AREA_NAME'],sql_clause=(None, 'ORDER BY CONTACT_ORGANIZATION_NAME, CNSLTN_AREA_NAME')):
    arcpy.AddMessage("{:50}{:5}{}".format(row[0],"|",row[1]))


arcpy.Delete_management('in_memory')
Environment.delete_bcgw_connection(location=sde_folder)
