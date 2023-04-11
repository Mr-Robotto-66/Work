import arcpy
import arcpyproduction
import os
import sys
import logging
import pandas as pd
import math

from argparse import ArgumentParser
from collections import defaultdict
from collections import OrderedDict
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
from util.cls_ogma_statistics import OGMAStatistics
from util.cls_ogma_targets import OGMATarget

sys.path.insert(1, r'W:\FOR\RSI\TOC\Projects\ESRI_Scripts\Python_Repository')

from environment import Environment
from erase_features import EraseFeatures
from excel import Excel
arcpy.env.outputMFlag = "Disabled"
arcpy.env.outputZFlag = "Disabled"

def run_app():
    tsa, out, un, pw, analyze, report, script_dir, logger = get_input_parameters()
    ogma = OgmaAnalysis(tsa=tsa, output_location=out, username=un, password=pw,
                        analyze=analyze, report=report, script_dir=script_dir, logger=logger)
    if ogma.analyze:
        ogma.prepare_data()
        ogma.create_aoi()
        ogma.identity_aoi()
        ogma.update_attributes()
    if ogma.report:
        ogma.build_statistics()
        ogma.create_report()
    del ogma


def get_input_parameters():
    try:
        parser = ArgumentParser(description='This script performs analysis on each landscape unit within the '
                                            'selected TSA to determine OGMA allocation.  It will output reports and '
                                            'maps which include operating areas and tfls within the selected '
                                            'landscape unit')
        parser.add_argument('tsa', type=str, help='Timber Supply Area Name')
        parser.add_argument('out', type=str, help='Output Location')
        parser.add_argument('un', type=str, help='BCGW username')
        parser.add_argument('pw', type=str, help='BCGW password')
        parser.add_argument('analyze', type=str, nargs='?', help='Analyze Data')
        parser.add_argument('report', type=str, nargs='?', help='Generate Report')
        parser.add_argument('--log_level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                            help='Log level')
        parser.add_argument('--log_dir', help='Path to log directory')

        args = parser.parse_args()

        logger = Environment.setup_logger(args)

        script_dir = os.path.dirname(sys.argv[0])

        return args.tsa, args.out, args.un, arcpy.GetParameterAsText(3), args.analyze, args.report, script_dir, logger

    except Exception as e:
        logging.error('Unexpected exception. Program terminating: {}'.format(e.message))
        raise Exception('Errors exist')


def get_value_from_range(num, lst_breaks, lst_results):
    if num == 0:
        return 0
    for i in range(1, len(lst_breaks)):
        if lst_breaks[i - 1] <= num <= lst_breaks[i]:
            return lst_results[i - 1]
        else:
            if i == len(lst_breaks) - 1:
                if num > lst_breaks[i]:
                    return lst_results[i]


class OgmaAnalysis:
    def __init__(self, tsa, output_location, username, password, analyze, report, script_dir, logger):
        # Assign parameters and workspace variables
        self.tsa = tsa
        self.out_dir = output_location
        self.sde_folder = 'Database Connections'
        self.script_dir = script_dir

        # self.str_lu_name = self.lu_name.split(':')[1].split('(')[0] \
        #     if ':' in self.lu_name else self.lu_name.split('(')[0]
        # self.str_lu_number = self.lu_name.split(':')[0] if ':' in self.lu_name else None
        # self.int_lu_id = int(self.lu_name.split(':')[1].split('(')[1][:-1]
        #                      if ':' in self.lu_name else self.lu_name.split('(')[1][:-1])

        self.data_dir = os.path.join(self.out_dir, self.tsa, 'Data')
        self.plot_dir = os.path.join(self.out_dir, self.tsa, 'Plots')
        self.report_dir = os.path.join(self.out_dir, self.tsa, 'Reports')
        self.out_gdb = os.path.join(self.data_dir, 'OGMA_Data.gdb')
        self.analyze = True if analyze.lower() == 'true' else False
        self.report = True if report.lower() == 'true' else False
        self.logger = logger

        # Connect to SDE databases and create output folders
        self.lrm_db = Environment.create_lrm_connection(location=self.sde_folder, lrm_user_name='map_view_14',
                                                        lrm_password='interface', logger=self.logger)

        self.bcgw_db = Environment.create_bcgw_connection(location=self.sde_folder, bcgw_user_name=username,
                                                          bcgw_password=password, logger=self.logger)

        self.lu_file = os.path.join(self.script_dir, 'templates', 'landscape_units.csv')
        self.target_file = os.path.join(self.script_dir, 'templates', 'ogma_targets.csv')
        self.mxd_template = os.path.join(self.script_dir, 'templates', 'ogma_lu_template.mxd')

        self.excel_report_file = os.path.join(self.report_dir, '{}_LU_OGMA_{}.xlsx'
                                              .format(self.tsa, dt.now().strftime('%Y%m%d')))

        self.str_ogma_summary_targets = ''
        self.str_ogma_age_class = ''

        self.bl_corridor = True if self.tsa == 'Golden' else False

        self.logger.info('Preparing workspace')
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        if not os.path.exists(self.plot_dir):
            os.makedirs(self.plot_dir)
        if not os.path.exists(self.report_dir):
            os.makedirs(self.report_dir)

        if not arcpy.Exists(self.out_gdb):
            arcpy.CreateFileGDB_management(out_folder_path=os.path.dirname(self.out_gdb),
                                           out_name=os.path.basename(self.out_gdb))

        # Field names
        self.fld_lu_name = 'LANDSCAPE_UNIT_NAME'
        self.fld_lu_number = 'LANDSCAPE_UNIT_NUMBER'
        self.fld_lu_id = 'LANDSCAPE_UNIT_PROVID'
        self.fld_lu_bio = 'BIODIVERSITY_EMPHASIS_OPTION'
        self.fld_nat_dist = 'NATURAL_DISTURBANCE'
        self.fld_zone = 'ZONE'
        self.fld_status = 'STATUS'
        self.fld_operable = 'OPERABLE'
        self.fld_age = 'AGE_UPDATED'
        self.fld_age_class = 'AGE_CLASS_UPDATED'
        self.fld_age_type = 'AGE_TYPE'
        self.fld_proj_age = 'PROJ_AGE_1'
        self.fld_proj_date = 'PROJECTED_DATE'
        self.fld_cc_status = 'CC_STATUS'
        self.fld_cc_harvest_date = 'CC_HARVEST_DATE'
        self.fld_land_type = 'LAND_TYPE'
        self.fld_bclcs_1 = 'BCLCS_LEVEL_1'
        self.fld_bclcs_2 = 'BCLCS_LEVEL_2'
        self.fld_bclcs_3 = 'BCLCS_LEVEL_3'
        self.fld_bclcs_4 = 'BCLCS_LEVEL_4'
        self.fld_fmlb_ind = 'FOR_MGMT_LAND_BASE_IND'
        self.fld_line_7b = 'LINE_7B_DISTURBANCE_HISTORY'
        self.fld_crown_closure = 'CROWN_CLOSURE'
        self.fld_line_activity = 'LINE_7_ACTIVITY_HIST_SYMBOL'
        self.fld_area = 'SHAPE@AREA'
        self.fld_lr_name = 'STRGC_LAND_RSRCE_PLAN_NAME'
        self.fld_op_area = 'OPERATING_AREA'
        self.fld_corridor = 'CORRIDOR'


        # Source data
        self.__landscape_unit = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_LANDSCAPE_UNIT_SVW')
        self.__business_area = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.FADM_BCTS_AREA_SP')
        self.__land_resource_plans = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_STRGC_LAND_RSRCE_PLAN_SVW')
        self.__timber_supply_areas = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.FADM_TSA')
        self.__operating_areas = os.path.join(self.lrm_db, 'BCTS_SPATIAL.BCTS_PROV_OP', 'BCTS_SPATIAL.OPERATING_AREA')

        self.__dict_source_data = {
            'private land': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_CADASTRE.PMBC_PARCEL_FABRIC_POLY_FA_SVW'),
                                      sql='OWNER_TYPE = \'Private\'', data_type='REMOVE'),
            'crown reversions': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_TANTALIS.TA_REVERSION_SHAPES')),
            'woodlots': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_MANAGED_LICENCE_POLY_SVW'),
                                  data_type='REMOVE'),
            # 'national parks': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.CLAB_NATIONAL_PARKS'),
            #                             data_type='REMOVE'),
            'provincial parks': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_TANTALIS.TA_PARK_ECORES_PA_SVW'),
                                          data_type='REMOVE'),
            'crown federal land':
                OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_CADASTRE.CBM_INTGD_CADASTRAL_FABRIC_SVW'),
                          sql='OWNERSHIP_CLASS = \'CROWN FEDERAL\'', data_type='REMOVE'),
            'vri': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.VEG_COMP_LYR_R1_POLY'),
                             data_type='ADD'),
            'bec': OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.BEC_BIOGEOCLIMATIC_POLY'),
                             data_type='ADD'),
            'toc ogma': OgmaInput(path=r'\\bctsdata.bcgov\data\toc_root\Local_Data\ogma\TOC_OGMA.shp',
                                  sql='Status <> \'D\''),
            'toc mogma': OgmaInput(path=r'\\bctsdata.bcgov\data\toc_root\Local_Data\ogma\TOC_MOGMA.shp',
                                   sql='Status IN (\'A\', \'MOGMA\')'),
            'provincial ogma':
                OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_OGMA_NON_LEGAL_CURRENT_SVW')),
            'consolidated cutblocks': OgmaInput(path=r'\\spatialfiles2.bcgov\Archive\FOR\RSI\TOC\Local_Data'
                                                     r'\Data_Library\forest\consolidated_cutblocks'
                                                     r'\consolidated_cutblocks.gdb\ConsolidatedCutblocks_Prod_Res',
                                                data_type='ADD'),
            'operating areas':
                OgmaInput(path=self.__operating_areas, data_type='ADD', sql='ORG_UNIT_CODE = \'TOC\''),
            'operability dos':
                OgmaInput(path=os.path.join(self.bcgw_db, 'REG_LAND_AND_NATURAL_RESOURCE.OPERABILITY_AREAS_SIR_POLY'),
                          sql='OPER = \'A\' OR OPER = \'H\''),
            'operability revelstoke':
                OgmaInput(path=os.path.join(self.bcgw_db, 'REG_LAND_AND_NATURAL_RESOURCE.OPERABILITY_TRV_POLY'),
                          sql='OCL2002 = \'A\''),
            'operability golden':
                OgmaInput(path=os.path.join(self.bcgw_db, 'REG_LAND_AND_NATURAL_RESOURCE.OPERABILITY_TGD_POLY'),
                          sql='OPER = \'A\''),
            'operability cascadia':
                OgmaInput(path=r'\\spatialfiles2.bcgov\Archive\FOR\RSI\TOC\Local_Data\Data_Library\operability'
                               r'\Operability.gdb\operability_cascadia', sql='OPER <> \'N\' AND OPER <> \'I\''),
            'connectivity corridors':
                OgmaInput(path=os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_PLAN_LEGAL_POLY_SVW'),
                          sql='STRGC_LAND_RSRCE_PLAN_NAME = \'Kootenay Boundary Higher Level Plan Order\' AND '
                              'LEGAL_FEAT_OBJECTIVE = \'Connectivity Corridors\' AND LEGAL_FEAT_ATRB_1_VALUE <> \'0\'',
                          data_type='ADD'),
            'slope':
                OgmaInput(path=r'\\spatialfiles2.bcgov\Archive\FOR\RSI\TOC\Local_Data\Data_Library\terrain\Slope'
                               r'\Slope80.gdb\Slope80_LiDAR_DEM_Merge_TSAOnly')
        }

        # Resultant data
        self.fc_lu = os.path.join(self.out_gdb, 'landscape_unit')
        self.fc_aoi = os.path.join(self.out_gdb, 'aoi')
        self.fc_resultant = os.path.join(self.out_gdb, 'resultant')
        self.fc_lr_plans = os.path.join(self.out_gdb, 'lr_plans')
        self.fc_beo = os.path.join(self.out_gdb, 'beo')
        self.fc_ogma = os.path.join(self.out_gdb, 'ogma')
        self.dict_resultant_data = defaultdict(OgmaInput)

        # Other Variables
        self.lst_bio_options = ['Low', 'Intermediate', 'High']
        self.lst_age_class_breaks = [0, 20, 40, 60, 80, 100, 120, 140, 250]
        self.lst_age_class = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        self.lst_lu_names = []
        self.lst_lu_numbers = []

        self.dict_resource_plans = {
            'Okanagan Shuswap Land and Resource Management Plan': 'OKANAGAN SHUSWAP',
            'Revelstoke Higher Level Plan Order': 'REVELSTOKE',
            'Kootenay Boundary Higher Level Plan Order': 'KOOTENAY BOUNDARY'
        }
        self.dict_age_class = {
            0: 'Harvested',
            1: '1 to 20',
            2: '21 to 40',
            3: '41 to 60',
            4: '61 to 80',
            5: '81 to 100',
            6: '101 to 120',
            7: '121 to 140',
            8: '141 to 250',
            9: '251 +'
        }

        self.ogma_statistics = None
        self.ogma_targets = None

        self.str_forest = 'FORESTED'
        self.str_reserve = 'RESERVE'
        self.str_np = 'NON-PRODUCTIVE'
        self.str_harvest = 'HARVESTED'
        self.str_operable = 'OPERABLE'
        self.str_outside_oa = 'Outside Operating Area'

    def __del__(self):
        Environment.delete_lrm_connection(location=self.sde_folder, logger=self.logger)
        Environment.delete_bcgw_connection(location=self.sde_folder, logger=self.logger)

    def prepare_data(self):

        self.logger.info('Extracting landscape units')
        lyr_tsa = arcpy.MakeFeatureLayer_management(in_features=self.__timber_supply_areas, out_layer='lyr_tsa',
                                                    where_clause='(TSA_NUMBER IN (\'22\', \'27\', \'07\') AND '
                                                                 'TSB_NUMBER IS NULL OR '
                                                                 'COMMENTS = \'Cascadia TSA Block 4\') AND '
                                                                 'TSA_NUMBER_DESCRIPTION = \'{} TSA\''.format(self.tsa))

        lu_where_clause = '{0} IN ({1})'.format(self.fld_lu_bio,
                                                ','.join('\'{0}\''.format(bio) for bio in self.lst_bio_options))

        lyr_lu = arcpy.MakeFeatureLayer_management(in_features=self.__landscape_unit, out_layer='lyr_lu',
                                                   where_clause=lu_where_clause)

        lyr_oa = arcpy.MakeFeatureLayer_management(in_features=self.__operating_areas, out_layer='lyr_oa',
                                                   where_clause='ORG_UNIT_CODE = \'TOC\'')

        arcpy.SelectLayerByLocation_management(in_layer=lyr_oa, overlap_type='HAVE_THEIR_CENTER_IN',
                                               select_features=lyr_tsa, selection_type='NEW_SELECTION')

        arcpy.SelectLayerByLocation_management(in_layer=lyr_lu, overlap_type='INTERSECT',
                                               select_features=lyr_oa, selection_type='NEW_SELECTION',
                                               search_distance='-1000 METERS')
        arcpy.SelectLayerByLocation_management(in_layer=lyr_lu, overlap_type='CONTAINS',
                                               select_features=lyr_oa, selection_type='ADD_TO_SELECTION')

        self.lst_lu_names = sorted(list(set([row[0] for row in arcpy.da.SearchCursor(lyr_lu, self.fld_lu_name)])))
        self.lst_lu_numbers = sorted(list(set(['{}P'.format(row[0])
                                               for row in arcpy.da.SearchCursor(lyr_lu, self.fld_lu_number)])))

        lu_where_clause = '({0} IN ({1}) AND {2} IN ({3})) OR ({0} = \'NA\' AND {4} IN ({5}))' \
            .format(self.fld_lu_bio, ','.join('\'{0}\''.format(bio) for bio in self.lst_bio_options), self.fld_lu_name,
                    ','.join('\'{0}\''.format(lu) for lu in self.lst_lu_names), self.fld_lu_number,
                    ','.join('\'{0}\''.format(num) for num in self.lst_lu_numbers))

        arcpy.Select_analysis(in_features=self.__landscape_unit, out_feature_class=self.fc_lu,
                              where_clause=lu_where_clause)

        arcpy.Delete_management(in_data=lyr_lu)

        lyr_lu = arcpy.MakeFeatureLayer_management(in_features=self.fc_lu, out_layer='lyr_lu')

        with arcpy.da.SearchCursor(lyr_tsa, 'SHAPE@') as s_cursor:
            for row in s_cursor:
                tsa_geom = row[0]

        with arcpy.da.UpdateCursor(self.fc_lu, 'SHAPE@') as u_cursor:
            for row in u_cursor:
                if tsa_geom.disjoint(row[0]):
                    u_cursor.deleteRow()

        arcpy.RecalculateFeatureClassExtent_management(in_features=self.fc_lu)

        arcpy.Delete_management(in_data=lyr_tsa)
        arcpy.Delete_management(in_data=lyr_lu)
        arcpy.Delete_management(in_data=lyr_oa)

        self.logger.info('Selecting out land resource plans')
        where_clause = '{0} IN ({1})'.format(self.fld_lr_name, ','.join(
            '\'{0}\''.format(lr) for lr in self.dict_resource_plans.keys()))

        arcpy.Select_analysis(in_features=self.__land_resource_plans, out_feature_class=self.fc_lr_plans,
                              where_clause=where_clause)

        arcpy.env.extent = arcpy.Describe(value=self.fc_lu).extent

        for src in self.__dict_source_data:
            if src in ['connectivity corridors', 'slope'] and not self.bl_corridor:
                continue
            self.logger.info('Copying {0}'.format(src))
            fc_out = os.path.join(self.out_gdb, src.replace(' ', '_'))
            if not self.__dict_source_data[src].sql:
                arcpy.CopyFeatures_management(in_features=self.__dict_source_data[src].path, out_feature_class=fc_out)
            else:
                arcpy.Select_analysis(in_features=self.__dict_source_data[src].path, out_feature_class=fc_out,
                                      where_clause=self.__dict_source_data[src].sql)
            self.dict_resultant_data[src].path = fc_out
            self.dict_resultant_data[src].data_type = self.__dict_source_data[src].data_type
            if src == 'bec':
                beo_temp = '{}_temp'.format(self.fc_beo)
                arcpy.Dissolve_management(in_features=fc_out, out_feature_class=beo_temp,
                                          dissolve_field=[self.fld_nat_dist, self.fld_zone], multi_part='SINGLE_PART')
                arcpy.Intersect_analysis(in_features=[beo_temp, self.fc_lu], out_feature_class=self.fc_beo,
                                         join_attributes='NO_FID')
                arcpy.Delete_management(in_data=beo_temp)


        self.logger.info('Combining OGMA and MOGMA')
        ogma_merge = os.path.join(self.out_gdb, 'ogma_merge')
        # ogma_dissolve = os.path.join(self.out_gdb, 'ogma')
        lst_ogmas = []
        for fc in self.dict_resultant_data:
            if 'ogma' in fc:
                lst_ogmas.append(self.dict_resultant_data[fc].path)
        lst_ogmas.sort(reverse=True)
        arcpy.Merge_management(inputs=lst_ogmas, output=ogma_merge)
        arcpy.Dissolve_management(in_features=ogma_merge, out_feature_class=self.fc_ogma, multi_part='SINGLE_PART')
        arcpy.AddField_management(in_table=self.fc_ogma, field_name=self.fld_status, field_type='TEXT',
                                  field_length=25)
        with arcpy.da.UpdateCursor(self.fc_ogma, self.fld_status) as u_cursor:
            for row in u_cursor:
                row[0] = 'OGMA'
                u_cursor.updateRow(row)

        for fc in lst_ogmas + [ogma_merge]:
            arcpy.Delete_management(in_data=fc)
            if fc in self.dict_resultant_data:
                del self.dict_resultant_data[fc]
        self.dict_resultant_data['ogma'].path = self.fc_ogma
        self.dict_resultant_data['ogma'].data_type = 'ADD'

        self.logger.info('Combining operability areas')
        oper_merge = os.path.join(self.out_gdb, 'oper_merge')
        oper_dissolve = os.path.join(self.out_gdb, 'operability')
        lst_oper = []
        for fc in self.dict_resultant_data:
            if 'operability' in fc:
                lst_oper.append(self.dict_resultant_data[fc].path)
        lst_oper.sort(reverse=True)
        arcpy.Merge_management(inputs=lst_oper, output=oper_merge)
        arcpy.Dissolve_management(in_features=oper_merge, out_feature_class=oper_dissolve, multi_part='SINGLE_PART')
        arcpy.AddField_management(in_table=oper_dissolve, field_name=self.fld_operable,
                                  field_type='TEXT', field_length=25)
        with arcpy.da.UpdateCursor(oper_dissolve, self.fld_operable) as u_cursor:
            for row in u_cursor:
                row[0] = self.str_operable
                u_cursor.updateRow(row)
        for fc in lst_oper + [oper_merge]:
            arcpy.Delete_management(in_data=fc)
            if fc in self.dict_resultant_data:
                del self.dict_resultant_data[fc]
        self.dict_resultant_data['operability'].path = oper_dissolve
        self.dict_resultant_data['operability'].data_type = 'ADD'

        self.logger.info('Removing reversions from private land')
        p_lyr = arcpy.MakeFeatureLayer_management(in_features=self.dict_resultant_data['private land'].path,
                                                  out_layer='p_lyr')
        r_lyr = arcpy.MakeFeatureLayer_management(in_features=self.dict_resultant_data['crown reversions'].path,
                                                  out_layer='r_lyr')

        arcpy.SelectLayerByLocation_management(in_layer=p_lyr, overlap_type='HAVE_THEIR_CENTER_IN',
                                               select_features=r_lyr, selection_type='NEW_SELECTION')

        arcpy.DeleteFeatures_management(in_features=p_lyr)
        arcpy.Delete_management(p_lyr)
        arcpy.Delete_management(r_lyr)

        if self.bl_corridor:
            conn_slope = os.path.join(self.out_gdb, 'conn_slope')
            e_obj = EraseFeatures(in_features=self.dict_resultant_data['connectivity corridors'].path,
                                  erase_features=self.dict_resultant_data['slope'].path,
                                  out_features=conn_slope, logger=self.logger, add_layer=False)
            e_obj.erase_analysis()
            del e_obj

            arcpy.AddField_management(in_table=conn_slope, field_name=self.fld_corridor, field_type='TEXT',
                                      field_length=5)
            with arcpy.da.UpdateCursor(conn_slope, self.fld_corridor) as u_cursor:
                for row in u_cursor:
                    row[0] = 'YES'
                    u_cursor.updateRow(row)

            self.dict_resultant_data['connectivity corridors'].path = conn_slope
            self.dict_resultant_data['connectivity corridors'].data_type = 'ADD'

    def create_aoi(self):
        self.logger.info('Creating aoi')
        temp_fc = os.path.join(self.out_gdb, 'temp_fc')

        arcpy.CopyFeatures_management(in_features=self.fc_lu, out_feature_class=self.fc_aoi)

        for fc in self.dict_resultant_data:
            if self.dict_resultant_data[fc].data_type == 'REMOVE':
                # self.logger.info('Removing {}'.format(fc))
                e_obj = EraseFeatures(in_features=self.fc_aoi, erase_features=self.dict_resultant_data[fc].path,
                                      out_features=temp_fc, logger=self.logger, add_layer=False)
                e_obj.erase_analysis()
                del e_obj
                arcpy.CopyFeatures_management(in_features=temp_fc, out_feature_class=self.fc_aoi)
        arcpy.Delete_management(in_data=temp_fc)

    def identity_aoi(self):
        self.logger.info('Adding features to aoi')
        temp_fc = os.path.join(self.out_gdb, 'temp_fc')

        arcpy.CopyFeatures_management(in_features=self.fc_aoi, out_feature_class=self.fc_resultant)

        for fc in self.dict_resultant_data:
            if self.dict_resultant_data[fc].data_type == 'ADD':
                self.logger.info('Adding {}'.format(fc))
                arcpy.Union_analysis(in_features=[self.fc_resultant, self.dict_resultant_data[fc].path],
                                     out_feature_class=temp_fc)
                arcpy.Select_analysis(in_features=temp_fc, out_feature_class=self.fc_resultant,
                                      where_clause='FID_{} <> -1'.format(os.path.basename(self.fc_resultant)))
                lst_del_fields = [field.name for field in arcpy.ListFields(self.fc_resultant)
                                  if field.name.startswith('FID_')]
                arcpy.Delete_management(in_data=temp_fc)
                arcpy.DeleteField_management(in_table=self.fc_resultant, drop_field=lst_del_fields)
        self.logger.info('Cleaning up slivers')
        result_lyr = arcpy.MakeFeatureLayer_management(in_features=self.fc_resultant, out_layer='result_lyr')
        arcpy.SelectLayerByAttribute_management(in_layer_or_view=result_lyr, selection_type='NEW_SELECTION',
                                                where_clause='Shape_Area <= 5')
        arcpy.Eliminate_management(in_features=result_lyr, out_feature_class=temp_fc)
        arcpy.CopyFeatures_management(in_features=temp_fc, out_feature_class=self.fc_resultant)
        with arcpy.da.UpdateCursor(self.fc_resultant, self.fld_nat_dist) as u_cursor:
            for row in u_cursor:
                if row[0] == '':
                    u_cursor.deleteRow()

        arcpy.Delete_management(in_data=temp_fc)
        arcpy.Delete_management(in_data=result_lyr)

    def update_attributes(self):
        self.logger.info('Updating age and age class attributes')
        arcpy.AddField_management(in_table=self.fc_resultant, field_name=self.fld_age, field_type='SHORT')
        arcpy.AddField_management(in_table=self.fc_resultant, field_name=self.fld_age_class, field_type='SHORT')
        arcpy.AddField_management(in_table=self.fc_resultant, field_name=self.fld_land_type, field_type='TEXT')
        lst_fields = [self.fld_proj_age, self.fld_proj_date, self.fld_age, self.fld_age_class,
                      self.fld_cc_status, self.fld_cc_harvest_date]
        str_lrp_name = ''

        with arcpy.da.UpdateCursor(self.fc_resultant, lst_fields) as u_cursor:
            for row in u_cursor:
                if row[lst_fields.index(self.fld_proj_age)]:
                    now = dt.now()
                    proj_age = int(row[lst_fields.index(self.fld_proj_age)])
                    proj_date = row[lst_fields.index(self.fld_proj_date)]
                    cc_status = row[lst_fields.index(self.fld_cc_status)]
                    cc_harvest_date = row[lst_fields.index(self.fld_cc_harvest_date)]
                    if cc_status not in ('', self.str_reserve) and cc_harvest_date:
                        try:
                            proj_date = dt.strptime(cc_harvest_date, '%Y-%m-%d')
                        except:
                            proj_date = dt.strptime(cc_harvest_date, '%m/%d/%Y')
                        proj_age = 0
                    date_diff = relativedelta(now, proj_date)
                    if date_diff.years < 0:
                        date_diff.years = 0
                        row[lst_fields.index(self.fld_age)] = date_diff.years
                    else:
                        row[lst_fields.index(self.fld_age)] = proj_age + date_diff.years
                    row[lst_fields.index(self.fld_age_class)] = get_value_from_range(
                        row[lst_fields.index(self.fld_age)],
                        self.lst_age_class_breaks,
                        self.lst_age_class)
                u_cursor.updateRow(row)

        self.logger.info('Determining land resource plan')
        lu_lyr = arcpy.MakeFeatureLayer_management(in_features=self.fc_aoi, out_layer='lu_lyr')
        lrp_lyr = arcpy.MakeFeatureLayer_management(in_features=self.fc_lr_plans, out_layer='lrp_lyr')

        arcpy.SelectLayerByLocation_management(in_layer=lrp_lyr, overlap_type='CONTAINS',
                                               select_features=lu_lyr, selection_type='NEW_SELECTION')
        lst_lr_names = [row[0] for row in arcpy.da.SearchCursor(lrp_lyr, self.fld_lr_name)]
        for lr in lst_lr_names:
            str_lrp_name = lr

        if self.fld_lr_name not in [field.name for field in arcpy.ListFields(self.fc_resultant)]:
            arcpy.AddField_management(in_table=self.fc_resultant, field_name=self.fld_lr_name,
                                      field_type='TEXT', field_length=75)

        self.logger.info('Updating land type attributes')
        str_np_query = '{0} = \'N\' OR ({0} = \'V\' AND {1} = \'N\')'.format(self.fld_bclcs_1, self.fld_bclcs_2)
        lst_cursor_fields = [self.fld_land_type, self.fld_bclcs_1, self.fld_bclcs_2, self.fld_bclcs_3,
                             self.fld_bclcs_4, self.fld_fmlb_ind, self.fld_line_7b, self.fld_crown_closure,
                             self.fld_cc_status, self.fld_age, self.fld_line_activity, self.fld_status,
                             self.fld_operable, self.fld_age_class, self.fld_lr_name]

        with arcpy.da.UpdateCursor(self.fc_resultant, lst_cursor_fields) as u_cursor:
            for row in u_cursor:
                index_ltype = lst_cursor_fields.index(self.fld_land_type)
                bclcs_1 = row[lst_cursor_fields.index(self.fld_bclcs_1)]
                bclcs_2 = row[lst_cursor_fields.index(self.fld_bclcs_2)]
                bclcs_3 = row[lst_cursor_fields.index(self.fld_bclcs_3)]
                bclcs_4 = row[lst_cursor_fields.index(self.fld_bclcs_4)]
                fmlb_ind = row[lst_cursor_fields.index(self.fld_fmlb_ind)]
                line7b = row[lst_cursor_fields.index(self.fld_line_7b)]
                crown_closure = row[lst_cursor_fields.index(self.fld_crown_closure)]
                age = row[lst_cursor_fields.index(self.fld_age)]
                cc_status = row[lst_cursor_fields.index(self.fld_cc_status)]
                line_activity = row[lst_cursor_fields.index(self.fld_line_activity)]
                status = row[lst_cursor_fields.index(self.fld_status)]
                operable = row[lst_cursor_fields.index(self.fld_operable)]

                # Extract harvested
                if (age == 0 and cc_status != self.str_reserve) or (not age and line_activity == '$'):
                    row[index_ltype] = self.str_harvest
                    row[lst_cursor_fields.index(self.fld_age)] = 0
                    row[lst_cursor_fields.index(self.fld_age_class)] = 0

                # Extract non productive
                elif bclcs_1 == 'N' or \
                        (bclcs_2 == 'N' and bclcs_4 not in ('ST', 'SL')) or \
                        (bclcs_2 == 'N' and bclcs_3 == 'W') or \
                        bclcs_3 == 'A' or \
                        (fmlb_ind == 'N' and (line7b and not line7b.startswith('L'))) or \
                        (bclcs_2 == 'T' and bclcs_3 == 'W') or \
                        (bclcs_4 in ('ST', 'SL') and not (line7b and not line7b.startswith('L'))):
                    row[index_ltype] = self.str_np

                # Extract forested
                elif age > 0 or cc_status == self.str_reserve:
                    row[index_ltype] = self.str_forest

                if status == '':
                    row[lst_cursor_fields.index(self.fld_status)] = 'NON-OGMA'
                if operable == '':
                    row[lst_cursor_fields.index(self.fld_operable)] = 'INOPERABLE'

                row[lst_cursor_fields.index(self.fld_lr_name)] = str_lrp_name

                u_cursor.updateRow(row)

        try:
            arcpy.AddField_management(in_table=self.fc_resultant, field_name=self.fld_age_type,
                                      field_type='TEXT', field_length=10)
        except Exception as e:
            pass

        if not self.ogma_targets:
            self.build_targets()

        self.logger.info('Calculating age class type')
        lst_fields = [self.fld_nat_dist, self.fld_zone, self.fld_lu_bio, self.fld_age_class,
                      self.fld_age_type, self.fld_land_type, self.fld_lu_number]
        lr_plan = self.dict_resource_plans[str_lrp_name]
        with arcpy.da.UpdateCursor(self.fc_resultant, lst_fields) as u_cursor:
            for row in u_cursor:
                mature_age_class = None
                old_age_class = None
                ac_type = None
                ndt = row[lst_fields.index(self.fld_nat_dist)]
                bec = row[lst_fields.index(self.fld_zone)]
                beo = str(row[lst_fields.index(self.fld_lu_bio)]).upper()
                if beo == 'NA':
                    beo = 'HIGH'
                ac = row[lst_fields.index(self.fld_age_class)]
                land_type = row[lst_fields.index(self.fld_land_type)]
                lu_number = row[lst_fields.index(self.fld_lu_number)]
                beo_target = self.ogma_targets.lr_plan[lr_plan].ndt[ndt].bec_zone[bec].bio_opt[beo]
                if beo_target.mature.age:
                    mature_age_class = get_value_from_range(num=beo_target.mature.age + 1,
                                                            lst_breaks=self.lst_age_class_breaks,
                                                            lst_results=self.lst_age_class)
                if self.tsa == 'Golden' and 'G27' not in lu_number:
                    mature_age_class = None

                if beo_target.old.age:
                    old_age_class = get_value_from_range(num=beo_target.old.age + 1,
                                                         lst_breaks=self.lst_age_class_breaks,
                                                         lst_results=self.lst_age_class)

                if land_type in [self.str_forest, self.str_harvest]:
                    if ac < 3:
                        ac_type = 'EARLY'
                    elif mature_age_class and (3 <= ac < mature_age_class):
                        ac_type = 'MID'
                    elif not mature_age_class and (3 <= ac < old_age_class):
                        ac_type = 'MID'
                    elif (mature_age_class and old_age_class) and (mature_age_class <= ac < old_age_class):
                        ac_type = 'MATURE'
                    elif old_age_class and ac >= old_age_class:
                        ac_type = 'OLD'
                    else:
                        ac_type = None

                row[lst_fields.index(self.fld_age_type)] = ac_type

                u_cursor.updateRow(row)

    def build_statistics(self):
        lst_fields = [self.fld_lu_name, self.fld_lu_number, self.fld_nat_dist, self.fld_zone, self.fld_lu_bio,
                      self.fld_land_type, self.fld_age_class, self.fld_operable, self.fld_status, self.fld_area,
                      self.fld_lr_name, self.fld_age_type, self.fld_op_area]
        if self.bl_corridor:
            lst_fields.append(self.fld_corridor)

        self.logger.info('Building statistics')
        self.ogma_statistics = defaultdict(OGMAStatistics)
        lst_parks = []

        with arcpy.da.SearchCursor(self.fc_resultant, lst_fields) as s_cursor:
            for row in s_cursor:
                lu_name = row[lst_fields.index(self.fld_lu_name)]
                lu_number = str(row[lst_fields.index(self.fld_lu_number)])
                nat_dist = str(row[lst_fields.index(self.fld_nat_dist)]).upper()
                zone = str(row[lst_fields.index(self.fld_zone)]).upper()
                lu_bio = str(row[lst_fields.index(self.fld_lu_bio)]).upper()
                status = row[lst_fields.index(self.fld_status)]
                age_class = row[lst_fields.index(self.fld_age_class)]
                land_type = row[lst_fields.index(self.fld_land_type)]
                operable = row[lst_fields.index(self.fld_operable)]
                area = row[lst_fields.index(self.fld_area)] / 10000
                lr_name = row[lst_fields.index(self.fld_lr_name)]
                ac_type = row[lst_fields.index(self.fld_age_type)]
                op_area = row[lst_fields.index(self.fld_op_area)]
                corridor = row[lst_fields.index(self.fld_corridor)] if self.bl_corridor else None

                op_area = self.str_outside_oa if not op_area else op_area

                if lu_number.endswith('P'):
                    if lu_number not in lst_parks:
                        lst_parks.append(lu_number)
                    continue

                if self.ogma_statistics[lu_name].lr_plan == '':
                    self.ogma_statistics[lu_name].lr_plan = lr_name

                if self.ogma_statistics[lu_name].lu_number == '':
                    self.ogma_statistics[lu_name].lu_number = lu_number

                if (age_class or age_class >= 0) and land_type in [self.str_forest, self.str_harvest]:
                    # if ac_type:
                    self.ogma_statistics[lu_name].nat_disturbance[nat_dist].zone[zone].bio_opt[lu_bio].status[
                        status].age_class[age_class].op_areas[op_area].land_type[land_type].operable[
                        operable].area += area
                    self.ogma_statistics[lu_name].nat_disturbance[nat_dist].zone[zone].bio_opt[lu_bio].status[
                        status].age_class[age_class].ac_type = ac_type
                    if corridor == 'YES':
                        self.ogma_statistics[lu_name].nat_disturbance[nat_dist].zone[zone].bio_opt[lu_bio].status[
                            status].age_class[age_class].op_areas[op_area].conn_area += area

        if len(lst_parks) > 0:
            for park in lst_parks:
                for lu in self.ogma_statistics:
                    if self.ogma_statistics[lu].lu_number == park[:-1]:
                        self.ogma_statistics[lu].lu_park = OGMAStatistics()

                        with arcpy.da.SearchCursor(self.fc_resultant, lst_fields,
                                                   '{} = \'{}\''.format(self.fld_lu_number, park)) as s_cursor:
                            for row in s_cursor:
                                lu_name = row[lst_fields.index(self.fld_lu_name)]
                                lu_number = str(row[lst_fields.index(self.fld_lu_number)])
                                nat_dist = str(row[lst_fields.index(self.fld_nat_dist)]).upper()
                                zone = str(row[lst_fields.index(self.fld_zone)]).upper()
                                lu_bio = str(row[lst_fields.index(self.fld_lu_bio)]).upper()
                                status = row[lst_fields.index(self.fld_status)]
                                age_class = row[lst_fields.index(self.fld_age_class)]
                                land_type = row[lst_fields.index(self.fld_land_type)]
                                operable = row[lst_fields.index(self.fld_operable)]
                                area = row[lst_fields.index(self.fld_area)] / 10000
                                ac_type = row[lst_fields.index(self.fld_age_type)]
                                op_area = row[lst_fields.index(self.fld_op_area)]
                                corridor = row[lst_fields.index(self.fld_corridor)] if self.bl_corridor else None

                                op_area = self.str_outside_oa if not op_area else op_area

                                if not self.ogma_statistics[lu].park_name:
                                    self.ogma_statistics[lu].park_name = lu_name

                                if not self.ogma_statistics[lu].park_number:
                                    self.ogma_statistics[lu].park_number = lu_number

                                if (age_class or age_class >= 0) and land_type in [self.str_forest, self.str_harvest]:
                                    # if ac_type:
                                    self.ogma_statistics[lu].nat_disturbance[nat_dist].zone[zone].bio_opt[
                                        lu_bio].status[status].age_class[age_class].op_areas[op_area].land_type[
                                        land_type].operable[operable].area += area
                                    self.ogma_statistics[lu].nat_disturbance[nat_dist].zone[zone].bio_opt[
                                        lu_bio].status[status].age_class[age_class].ac_type = ac_type
                                    if corridor == 'YES':
                                        self.ogma_statistics[lu].nat_disturbance[nat_dist].zone[zone].bio_opt[
                                            lu_bio].status[
                                            status].age_class[age_class].op_areas[op_area].conn_area += area

        for lu in self.ogma_statistics:
            self.ogma_statistics[lu].total()

        if not self.ogma_targets:
            self.build_targets()

    def build_targets(self):
        self.logger.info('Building targets')
        df = pd.read_csv(filepath_or_buffer=self.target_file, delimiter=',').fillna(value='')

        lst_rows = [list(row) for row in df.values]
        lst_columns = df.columns.tolist()
        i_lr_plan = lst_columns.index('LAND_RESOURCE_PLAN')
        i_ndt = lst_columns.index(self.fld_nat_dist)
        i_zone = lst_columns.index(self.fld_zone)
        i_beo = lst_columns.index(self.fld_lu_bio)
        i_mature = lst_columns.index('MATURE')
        i_old = lst_columns.index('OLD')
        i_target_mature = lst_columns.index('TARGET_MATURE_OLD')
        i_target_old = lst_columns.index('TARGET_OLD')

        self.ogma_targets = OGMATarget()

        for row in lst_rows:
            lr_plan = row[i_lr_plan]
            ndt = str(row[i_ndt]).upper()
            zone = str(row[i_zone]).upper()
            beo = str(row[i_beo]).upper()
            mature = int(row[i_mature]) if row[i_mature] != '' else None
            old = int(row[i_old]) if row[i_old] != '' else None
            target_mature = float(row[i_target_mature]) if row[i_target_mature] != '' else None
            target_old = float(row[i_target_old]) if row[i_target_old] != '' else None

            self.ogma_targets.lr_plan[lr_plan].ndt[ndt].bec_zone[zone].bio_opt[beo].mature.age = mature
            self.ogma_targets.lr_plan[lr_plan].ndt[ndt].bec_zone[zone].bio_opt[beo].mature.target = target_mature
            self.ogma_targets.lr_plan[lr_plan].ndt[ndt].bec_zone[zone].bio_opt[beo].old.age = old
            self.ogma_targets.lr_plan[lr_plan].ndt[ndt].bec_zone[zone].bio_opt[beo].old.target = target_old

    def create_report(self):
        self.logger.info('Generating report')

        xl = Excel()

        xl_hal_center = xl.xl_hal_center
        xl_hal_left = xl.xl_hal_left
        xl_hal_right = xl.xl_hal_right
        xl_val_center = xl.xl_val_center

        xl_med = xl.xl_med
        xl_thin = xl.xl_thin

        xl_double = xl.xl_double
        xl_continuous = xl.xl_continuous

        white_colour = (255, 255, 255)
        red_colour = (192, 0, 0)
        early_colour = (255, 255, 190)
        mid_colour = (215, 194, 158)
        mature_colour = (171, 205, 102)
        old_colour = (92, 137, 68)
        black_colour = (0, 0, 0)
        deficit_colour = (255, 0, 0)
        surplus_colour = (0, 176, 80)
        brown_colour = (148, 138, 84)
        gray_colour = (166, 166, 166)
        light_gray_colour = (217, 217, 217)
        green_colour = (216, 228, 188)

        dict_ac_colours = {
            'EARLY': early_colour,
            'MID': mid_colour,
            'MATURE': mature_colour,
            'OLD': old_colour,
            '': white_colour,
            None: white_colour
        }

        xl.add_workbook()
        # xl.delete_sheet(3)
        # xl.delete_sheet(2)

        self.lst_lu_names = sorted(self.ogma_statistics.keys())

        for str_lu_name in self.lst_lu_names:
            lu_statistics = self.ogma_statistics[str_lu_name]
            sheet_name = str_lu_name
            sheet_title = str_lu_name
            if lu_statistics.park_name:
                sheet_title = '{}/{}-{}'.format(sheet_name, lu_statistics.park_number, lu_statistics.park_name)
            self.logger.info('Adding {}'.format(sheet_title))
            if str_lu_name == self.lst_lu_names[0]:
                xl.rename_sheet(1, sheet_name)
            else:
                xl.add_sheet(sheet=sheet_name)
            xl.activate_sheet(sheet_name)

            style_title = xl.add_style('title', bold=True, size=14)
            style_subtitle = xl.add_style('sub_title', bold=True, size=12, r_border=xl_thin, l_border=xl_thin,
                                          t_border=xl_thin, b_border=xl_thin, h_align=xl_hal_center)
            style_redboldtext = xl.add_style('red_bold_text', bold=True, r_border=xl_thin, l_border=xl_thin,
                                             t_border=xl_thin, b_border=xl_thin, h_align=xl_hal_right,
                                             text_colour=red_colour)
            style_text = xl.add_style('text', r_border=xl_thin, l_border=xl_thin, t_border=xl_thin,
                                      b_border=xl_thin, h_align=xl_hal_center)
            style_number = xl.add_style('number', r_border=xl_thin, l_border=xl_thin, t_border=xl_thin,
                                        b_border=xl_thin, h_align=xl_hal_right, cell_format='#,##0.00')
            style_percent = xl.add_style('percent', r_border=xl_thin, l_border=xl_thin, t_border=xl_thin,
                                         b_border=xl_thin, h_align=xl_hal_right, cell_format='0.00%')
            style_redboldnumber = xl.add_style('red_bold_number', bold=True, r_border=xl_thin, l_border=xl_thin,
                                               t_border=xl_thin, b_border=xl_thin, h_align=xl_hal_right,
                                               cell_format='#,##0.00', text_colour=red_colour)
            style_redboldpercent = xl.add_style('red_bold_percent', bold=True, r_border=xl_thin, l_border=xl_thin,
                                                t_border=xl_thin, b_border=xl_thin, h_align=xl_hal_right,
                                                cell_format='0.00%', text_colour=red_colour)

            xl.change_all_cell_colour(colour=white_colour)

            incr_one = 1 if self.bl_corridor else 0
            incr_two = 2 if self.bl_corridor else 0
            incr_three = 3 if self.bl_corridor else 0
            incr_four = 4 if self.bl_corridor else 0
            incr_five = 5 if self.bl_corridor else 0

            i_row = 1
            i_col = 1

            ndt_col = 1
            bec_col = 2
            bio_col = 3
            stat_col = 4
            ac_col = 5
            area_col = 6
            per_col = 7
            corr_col = 7 + incr_one
            oa_area_col = 8 + incr_one
            oa_per_col = 9 + incr_one
            oa_op_col = 10 + incr_one
            oa_op_p_col = 11 + incr_one
            oa_corr_col = 11 + incr_two

            s_ndt_col = 15 + incr_two
            s_zone_col = 16 + incr_two
            s_bio_col = 17 + incr_two
            s_area_col = 18 + incr_two
            s_ogma_area_col = 19 + incr_two
            s_corr_area_col = 19 + incr_three
            s_mat_col = 20 + incr_three
            s_mat_p_col = 21 + incr_three
            s_mat_targ_col = 22 + incr_three
            s_mat_targ_ha_col = 23 + incr_three
            s_mat_p_m_col = 24 + incr_three
            s_mat_corr_col = 24 + incr_four
            s_old_col = 25 + incr_four
            s_old_p_col = 26 + incr_four
            s_old_targ_col = 27 + incr_four
            s_old_targ_ha_col = 28 + incr_four
            s_old_p_m_col = 29 + incr_four
            s_old_corr_col = 29 + incr_five

            xl.write_range(i_row=i_row, j_row=i_row, i_col=i_col, j_col=i_col + 5,
                           value='Landscape Unit: {}'.format(sheet_title), style_name=style_title)
            i_row += 1

            xl.write_range(i_row=i_row, j_row=i_row, i_col=i_col, j_col=i_col + 10,
                           value='Land Resource Plan: {}'.format(lu_statistics.lr_plan), style_name=style_title)

            i_row += 2
            xl.change_style(style_name=style_title, size=12, h_align=xl_hal_center, l_border=xl_thin, r_border=xl_thin,
                            l_style=xl_double, r_style=xl_double)
            xl.write_range(i_row=i_row, j_row=i_row, i_col=area_col, j_col=corr_col, value='Landscape Unit',
                           style_name=style_title)
            xl.write_range(i_row=i_row, j_row=i_row, i_col=oa_area_col, j_col=oa_corr_col, value='All Operating Areas',
                           style_name=style_title)
            xl.change_style(style_name=style_title, size=12, h_align=xl_hal_left, l_border=0, r_border=0,
                            l_style=xl_continuous, r_style=xl_continuous)
            i_row += 1

            xl.change_style(style_name=style_subtitle, colour=gray_colour)
            xl.write_cell(i_row=i_row, i_col=ndt_col, value='NDT', style_name=style_subtitle)
            xl.write_cell(i_row=i_row, i_col=bec_col, value='BEC Zone', style_name=style_subtitle)
            xl.write_cell(i_row=i_row, i_col=bio_col, value='BEO', style_name=style_subtitle)
            xl.write_cell(i_row=i_row, i_col=stat_col, value='Status', style_name=style_subtitle)
            xl.write_cell(i_row=i_row, i_col=ac_col, value='Age Class', style_name=style_subtitle)
            xl.change_style(style_name=style_subtitle, l_style=xl_double)
            xl.write_cell(i_row=i_row, i_col=area_col, value='Area (Ha)', style_name=style_subtitle)
            xl.change_style(style_name=style_subtitle, l_style=xl_continuous)
            xl.write_cell(i_row=i_row, i_col=per_col, value='% of Total', style_name=style_subtitle)
            if self.bl_corridor:
                xl.write_cell(i_row=i_row, i_col=corr_col, value='Corridor Area (Ha)', style_name=style_subtitle)
            xl.change_style(style_name=style_subtitle, l_style=xl_double)
            xl.write_cell(i_row=i_row, i_col=oa_area_col, value='Area (Ha)', style_name=style_subtitle)
            xl.change_style(style_name=style_subtitle, l_style=xl_continuous)
            xl.write_cell(i_row=i_row, i_col=oa_per_col, value='% of Total', style_name=style_subtitle)
            xl.write_cell(i_row=i_row, i_col=oa_op_p_col, value='% Operable', style_name=style_subtitle)
            if not self.bl_corridor:
                xl.change_style(style_name=style_subtitle, r_style=xl_double)
            xl.write_cell(i_row=i_row, i_col=oa_op_col, value='Operable Area (Ha)', style_name=style_subtitle)
            if self.bl_corridor:
                xl.change_style(style_name=style_subtitle, r_style=xl_double)
                xl.write_cell(i_row=i_row, i_col=oa_corr_col, value='Corridor Area (Ha)', style_name=style_subtitle)
            xl.change_style(style_name=style_subtitle, r_style=xl_continuous)

            i_subtitle_row = i_row
            i_summary_row = i_row - 1

            i_row += 1
            dict_ndt = self.ogma_statistics[str_lu_name].nat_disturbance
            lu_number = self.ogma_statistics[str_lu_name].lu_number
            dict_summary = {}
            dict_oa_summary = defaultdict(lambda: defaultdict(Summary))

            str_lrp = self.dict_resource_plans[lu_statistics.lr_plan]
            lst_ndt_bec_bio = []
            for ndt in sorted(dict_ndt.keys()):
                xl.write_range(i_row=i_row, j_row=i_row + dict_ndt[ndt].ac_count + (dict_ndt[ndt].bio_count - 1),
                               i_col=ndt_col, j_col=ndt_col,
                               value=ndt, style_name=style_text)
                dict_bec = dict_ndt[ndt].zone
                for bec in sorted(dict_bec.keys()):
                    xl.write_range(i_row=i_row, j_row=i_row + dict_bec[bec].ac_count + len(dict_bec[bec].bio_opt) - 1,
                                   i_col=bec_col, j_col=bec_col, value=bec, style_name=style_text)

                    dict_bio = dict_bec[bec].bio_opt
                    for bio in sorted(dict_bio.keys()):
                        percent_total = 0
                        oa_percent_total = 0
                        xl.write_range(i_row=i_row, j_row=i_row + dict_bio[bio].ac_count - 1, i_col=bio_col,
                                       j_col=bio_col, value=bio, style_name=style_text)
                        dict_stat = dict_bio[bio].status
                        if bio == 'NA':
                            beo_target = self.ogma_targets.lr_plan[str_lrp].ndt[ndt].bec_zone[bec].bio_opt['HIGH']
                        else:
                            beo_target = self.ogma_targets.lr_plan[str_lrp].ndt[ndt].bec_zone[bec].bio_opt[bio]
                        summary = Summary(ndt=ndt, bec=bec, beo=bio)
                        summary.area = dict_bio[bio].area
                        if self.tsa == 'Golden':
                            if str_lu_name == 'Moose':
                                summary.mat_old_target = beo_target.mature.target
                            else:
                                summary.mat_old_target = None
                        else:
                            summary.mat_old_target = beo_target.mature.target
                        summary.old_target = beo_target.old.target
                        if lu_number == 'R3' and bio.upper() == 'LOW' and summary.old_target:
                            summary.old_target = round(summary.old_target * 3)

                        ndt_bec_bio = (ndt, bec, bio)
                        total_oa_area = 0
                        total_corr_area = 0
                        total_oa_op_area = 0
                        total_oa_corr_area = 0

                        for stat in sorted(dict_stat.keys()):
                            colour = light_gray_colour
                            if stat == 'OGMA':
                                colour = white_colour

                            for st in [style_text, style_number, style_percent]:
                                xl.change_style(style_name=st, colour=colour, bold=False)

                            xl.write_range(i_row=i_row, j_row=i_row + dict_stat[stat].ac_count - 1, i_col=stat_col,
                                           j_col=stat_col, value=stat, style_name=style_text)
                            dict_ac = dict_stat[stat].age_class

                            for ac in sorted(dict_ac.keys()):
                                bio_use = bio
                                if bio == 'NA':
                                    bio_use = 'HIGH'
                                dict_op_areas = dict_ac[ac].op_areas
                                ac_type = dict_ac[ac].ac_type
                                area = 0
                                corr_area = 0
                                oa_area = 0
                                oa_op_area = 0
                                oa_corr_area = 0
                                for oa in sorted(dict_op_areas.keys()):
                                    dict_type = dict_op_areas[oa].land_type
                                    area += dict_type[self.str_forest].area
                                    corr_area += dict_op_areas[oa].conn_area
                                    summary.corr_area += dict_op_areas[oa].conn_area
                                    total_corr_area += dict_op_areas[oa].conn_area
                                    if oa != self.str_outside_oa:
                                        total_oa_area += dict_type[self.str_forest].area
                                        oa_area += dict_type[self.str_forest].area
                                        oa_corr_area += dict_op_areas[oa].conn_area
                                        oa_op_area += dict_type[self.str_forest].operable[self.str_operable].area
                                        total_oa_op_area += dict_type[self.str_forest].operable[self.str_operable].area
                                        total_oa_corr_area += dict_op_areas[oa].conn_area

                                        if any([summary.mat_old_target, summary.old_target]):
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].ndt = ndt
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].bec = bec
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].beo = bio_use
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].area += \
                                                dict_type[self.str_forest].area
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].corr_area += \
                                                dict_op_areas[oa].conn_area
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_target = \
                                                summary.mat_old_target
                                            dict_oa_summary[oa][(ndt, bec, bio_use)].old_target = summary.old_target
                                            if stat == 'OGMA':
                                                dict_oa_summary[oa][(ndt, bec, bio_use)].ogma_area += area
                                                if ac_type == 'MATURE':
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_area += \
                                                        dict_type[self.str_forest].area
                                                    if self.tsa == 'Golden':
                                                        if str_lu_name == 'Moose':
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area += \
                                                                corr_area
                                                        else:
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area = None
                                                    else:
                                                        dict_oa_summary[oa][
                                                            (ndt, bec, bio_use)].mat_old_corr_area = None
                                                elif ac_type == 'OLD':
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_area += \
                                                        dict_type[self.str_forest].area
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].old_area += \
                                                        dict_type[self.str_forest].area
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].old_corr_area += \
                                                        corr_area
                                                    if self.tsa == 'Golden':
                                                        if str_lu_name == 'Moose':
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area += \
                                                                corr_area
                                                        else:
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area = None
                                                    else:
                                                        dict_oa_summary[oa][
                                                            (ndt, bec, bio_use)].mat_old_corr_area = None
                                                else:
                                                    # if self.tsa in ['Revelstoke', 'Cascadia', 'Golden']:
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_area += area
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].old_area += area
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].old_corr_area += \
                                                        corr_area
                                                    if self.tsa == 'Golden':
                                                        if str_lu_name == 'Moose':
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area += \
                                                                corr_area
                                                        else:
                                                            dict_oa_summary[oa][
                                                                (ndt, bec, bio_use)].mat_old_corr_area = None
                                                    else:
                                                        dict_oa_summary[oa][
                                                            (ndt, bec, bio_use)].mat_old_corr_area = None
                                            try:
                                                dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_pct = \
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].mat_old_area / \
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].area
                                                dict_oa_summary[oa][(ndt, bec, bio_use)].old_pct = \
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].old_area / \
                                                    dict_oa_summary[oa][(ndt, bec, bio_use)].area
                                            except:
                                                pass

                                    if ac == 0:
                                        area += dict_type[self.str_harvest].area
                                        if oa != self.str_outside_oa:
                                            oa_area += dict_type[self.str_harvest].area
                                            total_oa_area += dict_type[self.str_harvest].area
                                            if any([summary.mat_old_target, summary.old_target]):
                                                dict_oa_summary[oa][(ndt, bec, bio_use)].area += \
                                                    dict_type[self.str_harvest].area

                                if stat == 'OGMA':
                                    # if mature_age_class and (mature_age_class <= ac < old_age_class):
                                    if ac_type == 'MATURE':
                                        for st in [style_text, style_number, style_percent]:
                                            xl.change_style(style_name=st, bold=True, colour=mature_colour)

                                        if ndt_bec_bio not in lst_ndt_bec_bio:
                                            lst_ndt_bec_bio.append(ndt_bec_bio)

                                    # elif old_age_class and ac >= old_age_class:
                                    elif ac_type == 'OLD':
                                        for st in [style_text, style_number, style_percent]:
                                            xl.change_style(style_name=st, bold=True, colour=old_colour)

                                        if ndt_bec_bio not in lst_ndt_bec_bio:
                                            lst_ndt_bec_bio.append(ndt_bec_bio)
                                    summary.ogma_area += area
                                    if ac_type == 'MATURE':
                                        summary.mat_old_area += area
                                        if self.tsa == 'Golden':
                                            if str_lu_name == 'Moose':
                                                summary.mat_old_corr_area += corr_area
                                            else:
                                                summary.mat_old_corr_area = None
                                        else:
                                            summary.mat_old_corr_area = None
                                    elif ac_type == 'OLD':
                                        summary.mat_old_area += area
                                        summary.old_area += area
                                        summary.old_corr_area += corr_area
                                        if self.tsa == 'Golden':
                                            if str_lu_name == 'Moose':
                                                summary.mat_old_corr_area += corr_area
                                            else:
                                                summary.mat_old_corr_area = None
                                        else:
                                            summary.mat_old_corr_area = None
                                    else:
                                        # if self.tsa in ['Revelstoke', 'Cascadia', 'Golden']:
                                        summary.mat_old_area += area
                                        summary.old_area += area
                                        summary.old_corr_area += corr_area
                                        if self.tsa == 'Golden':
                                            if str_lu_name == 'Moose':
                                                summary.mat_old_corr_area += corr_area
                                            else:
                                                summary.mat_old_corr_area = None
                                        else:
                                            summary.mat_old_corr_area = None

                                xl.change_style(style_name=style_text, colour=dict_ac_colours[ac_type])
                                xl.write_cell(i_row=i_row, i_col=ac_col, value=ac, style_name=style_text)
                                xl.change_style(style_name=style_text, colour=white_colour, bold=False)
                                xl.change_style(style_name=style_number, l_style=xl_double)
                                xl.write_cell(i_row=i_row, i_col=area_col, value=area, style_name=style_number)
                                xl.write_cell(i_row=i_row, i_col=per_col,
                                              value=area/dict_bio[bio].area, style_name=style_percent)
                                percent_total += (area/dict_bio[bio].area)
                                if self.bl_corridor:
                                    xl.change_style(style_name=style_number, l_style=xl_continuous)
                                    xl.write_cell(i_row=i_row, i_col=corr_col, value=corr_area, style_name=style_number)
                                    xl.change_style(style_name=style_number, l_style=xl_double)
                                xl.write_cell(i_row=i_row, i_col=oa_area_col, value=oa_area, style_name=style_number)
                                xl.change_style(style_name=style_number, l_style=xl_continuous)
                                xl.write_cell(i_row=i_row, i_col=oa_per_col,
                                              value=oa_area / dict_bio[bio].area, style_name=style_percent)
                                oa_percent_total += (oa_area / dict_bio[bio].area)
                                xl.write_cell(i_row=i_row, i_col=oa_op_col, value=oa_op_area, style_name=style_number)
                                if not self.bl_corridor:
                                    xl.change_style(style_name=style_percent, r_style=xl_double)
                                try:
                                    xl.write_cell(i_row=i_row, i_col=oa_op_p_col,
                                                  value=oa_op_area / oa_area,
                                                  style_name=style_percent)
                                except Exception as e:
                                    xl.write_cell(i_row=i_row, i_col=oa_op_p_col, value=0, style_name=style_percent)

                                if self.bl_corridor:
                                    xl.change_style(style_name=style_number, r_style=xl_double)
                                    xl.write_cell(i_row=i_row, i_col=oa_corr_col, value=oa_corr_area,
                                                  style_name=style_number)
                                    xl.change_style(style_name=style_number, r_style=xl_continuous)
                                xl.change_style(style_name=style_percent, r_style=xl_continuous)

                                i_row += 1

                        summary.mat_old_pct = summary.mat_old_area / summary.area
                        summary.old_pct = summary.old_area / summary.area
                        if any([summary.mat_old_target, summary.old_target]):
                            if bio == 'NA':
                                dict_summary[(ndt, bec, 'HIGH')] + summary
                            else:
                                dict_summary[(ndt, bec, bio)] = summary
                            # lst_summary.append(summary)

                        xl.write_range(i_row=i_row, j_row=i_row, i_col=bio_col, j_col=ac_col,
                                       value='Sum ({} {} {})'.format(ndt, bec, bio), style_name=style_redboldtext)
                        xl.change_style(style_name=style_redboldnumber, l_style=xl_double)
                        xl.write_cell(i_row=i_row, i_col=area_col, value=dict_bio[bio].area,
                                      style_name=style_redboldnumber)
                        if not self.bl_corridor:
                            xl.change_style(style_name=style_redboldnumber, l_style=xl_double)
                        xl.write_cell(i_row=i_row, i_col=per_col, value=percent_total, style_name=style_redboldpercent)
                        if self.bl_corridor:
                            xl.write_cell(i_row=i_row, i_col=corr_col, value=total_corr_area,
                                          style_name=style_redboldnumber)
                            xl.change_style(style_name=style_redboldnumber, l_style=xl_double)
                        xl.write_cell(i_row=i_row, i_col=oa_area_col, value=total_oa_area,
                                      style_name=style_redboldnumber)
                        xl.change_style(style_name=style_redboldnumber, l_style=xl_continuous)
                        xl.write_cell(i_row=i_row, i_col=oa_op_col, value=total_oa_op_area,
                                      style_name=style_redboldnumber)
                        xl.write_cell(i_row=i_row, i_col=oa_op_p_col, value=total_oa_op_area / dict_bio[bio].area,
                                      style_name=style_redboldpercent)
                        if not self.bl_corridor:
                            xl.change_style(style_name=style_redboldpercent, r_style=xl_double)
                        xl.write_cell(i_row=i_row, i_col=oa_per_col, value=oa_percent_total,
                                      style_name=style_redboldpercent)
                        if self.bl_corridor:
                            xl.change_style(style_name=style_redboldnumber, r_style=xl_double)
                            xl.write_cell(i_row=i_row, i_col=oa_corr_col, value=total_oa_corr_area,
                                          style_name=style_redboldnumber)
                            xl.change_style(style_name=style_redboldnumber, r_style=xl_continuous)
                        xl.change_style(style_name=style_redboldpercent, r_style=xl_continuous)
                        i_row += 1

            def write_summary(title, row, summary_list):
                xl.change_style(style_name=style_title, size=12)
                xl.write_range(i_row=row - 1, j_row=row - 1, i_col=s_ndt_col, j_col=s_ndt_col + 6,
                               value=title, style_name=style_title)
                xl.change_style(style_name=style_subtitle, colour=green_colour)
                xl.write_cell(i_row=row, i_col=s_ndt_col, value='NDT', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_zone_col, value='BEC Zone', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_bio_col, value='BEO', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_area_col, value='Area (ha)', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_ogma_area_col, value='OGMA Area (ha)',
                              style_name=style_subtitle)
                if self.bl_corridor:
                    xl.write_cell(i_row=row, i_col=s_corr_area_col, value='Corridor Area (ha)',
                                  style_name=style_subtitle)
                xl.change_style(style_name=style_subtitle, colour=mature_colour, l_style=xl_double)
                xl.write_cell(i_row=row, i_col=s_mat_col, value='Mature+Old (ha)', style_name=style_subtitle)
                xl.change_style(style_name=style_subtitle, l_style=xl_continuous)
                xl.write_cell(i_row=row, i_col=s_mat_p_col, value='Mature+Old (%)', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_mat_targ_col, value='Target', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_mat_targ_ha_col, value='Target (ha)',
                              style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_mat_p_m_col, value='+/- (ha)', style_name=style_subtitle)
                if self.bl_corridor:
                    xl.write_cell(i_row=row, i_col=s_mat_corr_col, value='Corridor Area (ha)',
                                  style_name=style_subtitle)
                xl.change_style(style_name=style_subtitle, colour=old_colour, l_style=xl_double)
                xl.write_cell(i_row=row, i_col=s_old_col, value='Old (ha)', style_name=style_subtitle)
                xl.change_style(style_name=style_subtitle, l_style=xl_continuous)
                xl.write_cell(i_row=row, i_col=s_old_p_col, value='Old (%)', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_old_targ_col, value='Target', style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_old_targ_ha_col, value='Target (ha)',
                              style_name=style_subtitle)
                xl.write_cell(i_row=row, i_col=s_old_p_m_col, value='+/- (ha)', style_name=style_subtitle)
                if self.bl_corridor:
                    xl.write_cell(i_row=row, i_col=s_old_corr_col, value='Corridor Area (ha)',
                                  style_name=style_subtitle)
                for t in [style_text, style_number, style_percent]:
                    xl.change_style(style_name=t, colour=white_colour)

                for sm in summary_list:
                    row += 1

                    for t in [style_text, style_number, style_percent]:
                        xl.change_style(style_name=t, text_colour=black_colour)

                    xl.write_cell(i_row=row, i_col=s_ndt_col, value=sm.ndt, style_name=style_text)
                    xl.write_cell(i_row=row, i_col=s_zone_col, value=sm.bec, style_name=style_text)
                    xl.write_cell(i_row=row, i_col=s_bio_col, value=sm.beo, style_name=style_text)
                    xl.write_cell(i_row=row, i_col=s_area_col, value=sm.area, style_name=style_number)
                    xl.write_cell(i_row=row, i_col=s_ogma_area_col, value=sm.ogma_area,
                                  style_name=style_number)
                    if self.bl_corridor:
                        xl.write_cell(i_row=row, i_col=s_corr_area_col, value=sm.corr_area,
                                      style_name=style_number)
                    xl.change_style(style_name=style_number, l_style=xl_double)
                    xl.write_cell(i_row=row, i_col=s_mat_col,
                                  value=sm.mat_old_area if sm.mat_old_target else 'N/A', style_name=style_number)
                    xl.change_style(style_name=style_number, l_style=xl_continuous)
                    xl.write_cell(i_row=row, i_col=s_mat_targ_col,
                                  value='>{}'.format(sm.mat_old_target) if sm.mat_old_target else 'N/A',
                                  style_name=style_text)
                    target_ha = sm.area * (sm.mat_old_target / 100) if sm.mat_old_target else 'N/A'
                    target_plus_minus = sm.mat_old_area - target_ha if sm.mat_old_target else 'N/A'
                    xl.write_cell(i_row=row, i_col=s_mat_targ_ha_col, value=target_ha,
                                  style_name=style_number)

                    if (sm.mat_old_pct * 100) <= sm.mat_old_target and sm.mat_old_target:
                        xl.change_style(style_name=style_percent, text_colour=deficit_colour)
                        xl.change_style(style_name=style_number, text_colour=deficit_colour)
                    elif (sm.mat_old_pct * 100) > sm.mat_old_target and sm.mat_old_target:
                        xl.change_style(style_name=style_percent, text_colour=surplus_colour)
                        xl.change_style(style_name=style_number, text_colour=surplus_colour)
                    else:
                        xl.change_style(style_name=style_percent, text_colour=black_colour)
                        xl.change_style(style_name=style_number, text_colour=black_colour)

                    xl.write_cell(i_row=row, i_col=s_mat_p_col,
                                  value=sm.mat_old_pct if sm.mat_old_target else 'N/A', style_name=style_percent)
                    xl.write_cell(i_row=row, i_col=s_mat_p_m_col, value=target_plus_minus,
                                  style_name=style_number)

                    xl.change_style(style_name=style_percent, text_colour=black_colour)
                    xl.change_style(style_name=style_number, text_colour=black_colour)
                    if self.bl_corridor:
                        xl.write_cell(i_row=row, i_col=s_mat_corr_col,
                                      value=sm.mat_old_corr_area if sm.mat_old_corr_area else 'N/A',
                                      style_name=style_number)
                    xl.change_style(style_name=style_number, l_style=xl_double)
                    xl.write_cell(i_row=row, i_col=s_old_col,
                                  value=sm.old_area if sm.old_target else 'N/A', style_name=style_number)
                    xl.change_style(style_name=style_number, l_style=xl_continuous)
                    xl.write_cell(i_row=row, i_col=s_old_targ_col,
                                  value='>{}'.format(sm.old_target) if sm.old_target else 'N/A', style_name=style_text)
                    target_ha = sm.area * (sm.old_target / 100) if sm.old_target else 'N/A'
                    target_plus_minus = sm.old_area - target_ha if sm.old_target else 'N/A'
                    xl.write_cell(i_row=row, i_col=s_old_targ_ha_col, value=target_ha,
                                  style_name=style_number)

                    if (sm.old_pct * 100) <= sm.old_target and sm.old_target:
                        xl.change_style(style_name=style_percent, text_colour=deficit_colour)
                        xl.change_style(style_name=style_number, text_colour=deficit_colour)
                    elif (sm.old_pct * 100) > sm.old_target and sm.old_target:
                        xl.change_style(style_name=style_percent, text_colour=surplus_colour)
                        xl.change_style(style_name=style_number, text_colour=surplus_colour)
                    else:
                        xl.change_style(style_name=style_percent, text_colour=black_colour)
                        xl.change_style(style_name=style_number, text_colour=black_colour)

                    xl.write_cell(i_row=row, i_col=s_old_p_col,
                                  value=sm.old_pct if sm.old_target else 'N/A', style_name=style_percent)
                    xl.write_cell(i_row=row, i_col=s_old_p_m_col, value=target_plus_minus,
                                  style_name=style_number)
                    if self.bl_corridor:
                        xl.change_style(style_name=style_number, text_colour=black_colour)
                        xl.write_cell(i_row=row, i_col=s_old_corr_col, value=sm.old_corr_area,
                                      style_name=style_number)

                return row

            xl.write_range(i_row=i_summary_row, j_row=i_summary_row, i_col=s_ndt_col, j_col=s_ndt_col + 5,
                           value='Definition of Mature & Old Forests by NDT and Biogeoclimatic Zones',
                           style_name=style_title)
            i_summary_row += 1
            xl.change_style(style_name=style_subtitle, colour=brown_colour)
            xl.write_cell(i_row=i_summary_row, i_col=s_ndt_col, value='NDT', style_name=style_subtitle)
            xl.write_cell(i_row=i_summary_row, i_col=s_zone_col, value='BEC Zone', style_name=style_subtitle)
            xl.write_cell(i_row=i_summary_row, i_col=s_bio_col, value='Mature (yrs)', style_name=style_subtitle)
            xl.write_cell(i_row=i_summary_row, i_col=s_area_col, value='Old (yrs)', style_name=style_subtitle)

            i_summary_row += 1
            lst_ndt_bec = []
            for s in sorted(lst_ndt_bec_bio):
                ndt = s[0]
                bec = s[1]
                bio = s[2]
                if (ndt, bec) not in lst_ndt_bec:
                    age_targets = self.ogma_targets.lr_plan[str_lrp].ndt[ndt].bec_zone[bec].bio_opt[bio]
                    mat_age = '>{}'.format(age_targets.mature.age) if age_targets.mature.age else 'N/A'
                    old_age = '>{}'.format(age_targets.old.age) if age_targets.old.age else 'N/A'
                    xl.write_cell(i_row=i_summary_row, i_col=s_ndt_col, value=ndt, style_name=style_text)
                    xl.write_cell(i_row=i_summary_row, i_col=s_zone_col, value=bec, style_name=style_text)
                    xl.write_cell(i_row=i_summary_row, i_col=s_bio_col, value=mat_age, style_name=style_text)
                    xl.write_cell(i_row=i_summary_row, i_col=s_area_col, value=old_age, style_name=style_text)
                    lst_ndt_bec.append((ndt, bec))
                    i_summary_row += 1

            i_summary_row += 2

            i_summary_row = write_summary(title='OGMA LU Summary with Biodiversity Emphasis', row=i_summary_row,
                                          summary_list=[dict_summary[s] for s in sorted(dict_summary.keys())])

            i_summary_end_row = i_summary_row

            for oa in sorted(dict_oa_summary.keys()):
                if oa != self.str_outside_oa:
                    i_summary_row += 3
                    i_summary_row = write_summary(title='OGMA {} Summary with Biodiversity Emphasis'.format(oa),
                                                  row=i_summary_row,
                                                  summary_list=[dict_oa_summary[oa][s]
                                                                for s in sorted(dict_oa_summary[oa].keys())])

            i_summary_row += 2

            xl.change_style(style_name=style_text, colour=white_colour)
            xl.write_range(i_row=i_summary_row, j_row=i_summary_row, i_col=s_ndt_col, j_col=s_ndt_col + 2,
                           value='Age Classes', style_name=style_title)
            i_summary_row += 1
            for age in sorted(self.dict_age_class):
                if age == 0:
                    str_age_class = self.dict_age_class[age]
                else:
                    str_age_class = 'Stand age {}'.format(self.dict_age_class[age])
                xl.change_style(style_name=style_text, h_align=xl_hal_right)
                xl.write_cell(i_row=i_summary_row, i_col=s_ndt_col, value=age, style_name=style_text)
                xl.change_style(style_name=style_text, h_align=xl_hal_left)
                xl.write_range(i_row=i_summary_row, j_row=i_summary_row, i_col=s_ndt_col + 1, j_col=s_ndt_col + 2,
                               value=str_age_class, style_name=style_text)
                i_summary_row += 1

            xl.autofit_columns(start_col=1, end_col=s_old_corr_col, start_row=i_subtitle_row, end_row=i_row)

            self.str_ogma_age_class = os.path.join(os.path.dirname(self.excel_report_file), 'ogma_age_class.png')
            self.str_ogma_summary_targets = os.path.join(os.path.dirname(self.excel_report_file),
                                                         'ogma_summary_targets.png')

            xl.select_range(i_row=i_subtitle_row - 1, j_row=i_row - 1, i_col=1, j_col=oa_corr_col)
            xl.export_range(self.str_ogma_age_class)

            xl.select_range(i_row=i_subtitle_row - 1, j_row=i_summary_end_row, i_col=s_ndt_col, j_col=s_old_corr_col)
            xl.export_range(self.str_ogma_summary_targets)
            xl.save_workbook(file_path=self.excel_report_file)

            if lu_statistics.park_number:
                self.create_map(str_lu_name=str_lu_name, str_park_number=lu_statistics.park_number)
            else:
                self.create_map(str_lu_name=str_lu_name)
            os.remove(self.str_ogma_summary_targets)
            os.remove(self.str_ogma_age_class)

        xl.activate_sheet(self.lst_lu_names[0])
        xl.close_workbook(save=True, file_path=self.excel_report_file)
        xl.quit()
        # del xl

    def create_map(self, str_lu_name, str_park_number=None):
        from arcpy import mapping as mp

        self.logger.info('Creating map for {}'.format(str_lu_name))

        mxd = mp.MapDocument(self.mxd_template)
        df = mp.ListDataFrames(mxd)[0]

        str_add_query = ''
        if str_park_number:
            str_add_query = ' OR {} = \'{}\''.format(self.fld_lu_number, str_park_number)

        for elm in mp.ListLayoutElements(map_document=mxd):
            if elm.name == 'TITLE':
                if not str_park_number:
                    elm.text = '{}\nLU OGMA Analysis'.format(str_lu_name)
                else:
                    elm.text = '{}/{}\nLU OGMA Analysis'.format(str_lu_name, str_park_number)
                elm.elementPositionY += 0.1
            elif elm.name == 'ogma_age_class':
                elm.sourceImage = self.str_ogma_age_class
            elif elm.name == 'ogma_summary_targets':
                elm.sourceImage = self.str_ogma_summary_targets

            # if str_park_number:
            #     if elm.name in ('SUBTITLE', 'DETAILS', 'NAVIGATION'):
            #         elm.elementPositionY -= 0.1

        for lyr in mp.ListLayers(map_document_or_layer=mxd, data_frame=df):
            if lyr.name == 'Landscape Units':
                lyr.replaceDataSource(os.path.dirname(self.fc_lu), 'FILEGDB_WORKSPACE', os.path.basename(self.fc_lu))
                lyr.definitionQuery = '{} = \'{}\'{}'.format(self.fld_lu_name, str_lu_name, str_add_query)
            elif lyr.name == 'Seral Stage':
                lyr.replaceDataSource(os.path.dirname(self.fc_resultant), 'FILEGDB_WORKSPACE',
                                      os.path.basename(self.fc_resultant))
                lyr.definitionQuery = '{} IN (\'{}\', \'{}\') AND ({} =\'{}\'{})'.format(self.fld_land_type,
                                                                                         self.str_forest,
                                                                                         self.str_harvest,
                                                                                         self.fld_lu_name, str_lu_name,
                                                                                         str_add_query)
            elif lyr.name == 'Non-Productive':
                lyr.replaceDataSource(os.path.dirname(self.fc_resultant), 'FILEGDB_WORKSPACE',
                                      os.path.basename(self.fc_resultant))
                lyr.definitionQuery = '{} IN (\'{}\') AND ({} =\'{}\'{})'.format(self.fld_land_type, self.str_np,
                                                                                 self.fld_lu_name, str_lu_name,
                                                                                 str_add_query)
            elif lyr.name == 'BEC Zones':
                lyr.replaceDataSource(os.path.dirname(self.fc_beo), 'FILEGDB_WORKSPACE', os.path.basename(self.fc_beo))
            elif lyr.name == 'OGMA':
                lyr.replaceDataSource(os.path.dirname(self.fc_ogma), 'FILEGDB_WORKSPACE',
                                      os.path.basename(self.fc_ogma))
            elif lyr.name == 'Connectivity Corridors':
                if self.bl_corridor:
                    lyr.visible = True
                else:
                    lyr.visible = False

        full_geom = None
        with arcpy.da.SearchCursor(self.fc_lu, ['SHAPE@'],
                                   '{} = \'{}\'{}'.format(self.fld_lu_name, str_lu_name, str_add_query)) as s_cursor:
            for row in s_cursor:
                if not full_geom:
                    full_geom = row[0]
                else:
                    full_geom = full_geom.union(row[0])

        # df.scale = 50000
        df.extent = full_geom.extent
        arcpy.RefreshActiveView()
        df.scale = math.ceil(df.scale / 5000) * 5000

        arcpy.CheckOutExtension('Foundation')
        arcpyproduction.mapping.ClipDataFrameToGeometry(data_frame=df, clip_geometry=full_geom)
        arcpy.CheckInExtension('Foundation')

        self.logger.info('Exporting to pdf')
        pdf_map_file = os.path.join(self.plot_dir, '{}_LU_OGMA_{}.pdf'
                                    .format(str_lu_name, dt.now().strftime('%Y%m%d')))
        mp.ExportToPDF(map_document=mxd, out_pdf=pdf_map_file, image_quality='BETTER', image_compression='JPEG')
        # mxd.saveACopy(file_name='{}mxd'.format(pdf_map_file[:-3]))
        del mxd


class OgmaInput:
    def __init__(self, path=None, sql=None, data_type=None):
        self.path = path
        self.sql = sql
        self.data_type = data_type


class Summary:
    def __init__(self, ndt='', bec='', beo=''):
        self.ndt = ndt
        self.bec = bec
        self.beo = beo
        self.area = 0
        self.ogma_area = 0
        self.corr_area = 0
        self.mat_old_area = 0
        self.mat_old_pct = 0
        self.mat_old_target = 0
        self.mat_old_corr_area = 0
        self.old_area = 0
        self.old_pct = 0
        self.old_target = 0
        self.old_corr_area = 0

    def __add__(self, other):
        self.area += other.area
        self.ogma_area += other.ogma_area
        self.corr_area += other.corr_area
        self.mat_old_area += other.mat_old_area
        self.mat_old_corr_area += other.mat_old_corr_area
        self.old_area += other.old_area
        self.old_corr_area += other.old_corr_area
        self.mat_old_pct = self.mat_old_area / self.area
        self.old_pct = self.old_area / self.area



if __name__ == '__main__':
    run_app()
