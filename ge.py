import sys
import pprint
import xmltodict
import configparser
import pandas as pd
import pathlib
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

def extract_meta_from_config(file_name):
    # Exact meta data from config file (pca)

    try:
        with open(file_name, encoding='latin-1') as f:
            config = configparser.ConfigParser(interpolation=None)
            config.read_file(f)   

        sections = config.sections()
        meta_dict = {i: {i[0]: i[1] for i in config.items(i)} for i in config.sections()}
        # for key in meta_dict:
        #     print(key, meta_dict[key])
    except FileNotFoundError:
        print("ERROR: %s is missing. Looking for a _rar file" % file_name)
        fname        = pathlib.Path(file_name)
        fname_suffix = fname.suffix
        file_name_rar  = fname.with_name(file_name.stem + '_rar').with_suffix(fname_suffix)
        print('Found %s file' % file_name_rar)
        try:
            with open(file_name_rar, encoding='latin-1') as f:
                config = configparser.ConfigParser()
                config.read_file(f)   

            sections = config.sections()
            meta_dict = {i: {i[0]: i[1] for i in config.items(i)} for i in config.sections()}
        except:
            print("ERROR: %s is also missing" % file_name_rar)
            exit()
    return meta_dict

def extract_meta_from_dtxml(file_name):
    # Extract meta data from dtxml file
    with open(file_name, encoding='latin-1') as f:
        xml_content= f.read()

    xml_dict=xmltodict.parse(xml_content)
    i = 0
    meta_dict = {}
    while i < len(xml_dict['project']['additional_project_info']['property']):
        meta_dict[(xml_dict['project']['additional_project_info']['property'][i]['@name'])] =  xml_dict['project']['additional_project_info']['property'][i]['@value']
        i = i + 1

    # pprint.pprint(meta_dict)
    return meta_dict

def extract_meta_from_pcp(file_name):
    # Extract meta data from pcp file
    try:
        with open(file_name) as file:
            time_start = file.readlines()[2].split("\t")[-1].strip()
            datetime_start = datetime.strptime(time_start, '%Y-%m-%d %H:%M:%S')
            file.seek(0)  
            time_end = file.readlines()[-1].split("\t")[-1].strip()
            datetime_end = datetime.strptime(time_end, '%Y-%m-%d %H:%M:%S')
        scan_time = datetime_end - datetime_start
    except FileNotFoundError:
        print("ERROR: %s is missing. Looking for a _rar file" % file_name)
        fname        = pathlib.Path(file_name)
        fname_suffix = fname.suffix
        file_name_rar  = fname.with_name(file_name.stem + '_rar').with_suffix(fname_suffix)
        print('Found %s file' % file_name_rar)
        try:
            with open(file_name_rar, encoding='latin-1') as file:

                time_start = file.readlines()[2].split("\t")[-1].strip()
                datetime_start = datetime.strptime(time_start, '%Y-%m-%d %H:%M:%S')
                file.seek(0)  
                time_end = file.readlines()[-1].split("\t")[-1].strip()
                datetime_end = datetime.strptime(time_end, '%Y-%m-%d %H:%M:%S')
            scan_time = datetime_end - datetime_start
        except:
            print("ERROR: %s is also missing" % file_name_rar)
            exit()

    return scan_time, datetime_start

def main(args):

    if len(sys.argv) == 1:
        print ('ERROR: Must provide the path to a run-file folder as the argument')
        print ('Example:')
        print ('        python ge.py /Users/decarlo/conda/nocturn/data/FEG230530_413/')
        sys.exit(1)
    else:

        file_name   = sys.argv[1]
        p = pathlib.Path(file_name)
        if p.is_dir():
            p = pathlib.Path(file_name).joinpath(p.stem)
            file_name_pca   = p.with_suffix('.pca')
            file_name_pcp   = p.with_suffix('.pcp')
            file_name_pcr   = p.with_suffix('.pcr')
            file_name_dtxml = p.with_suffix('.dtxml')
            file_name_xlsx  = p.parents[1].joinpath('master').with_suffix('.xlsx')
        else:
            print('ERROR: %s does not exist' % p)
            sys.exit(1)
    # print('1', p)
    # print('2', p.suffix)
    # print('3', p.parents[0])
    # print('3', p.parents[1])
    # print('4', p.stem)
    # print('5', p.suffix)

    # print(file_name_pca  ) 
    # print(file_name_pcp  ) 
    # print(file_name_pcr  ) 
    # print(file_name_dtxml) 
    # print(file_name_xlsx ) 

    my_dict = {}

    scan_time, datetime_start = extract_meta_from_pcp(file_name_pcp)
    my_pcr_dict = extract_meta_from_config(file_name_pcr)
    my_pca_dict = extract_meta_from_config(file_name_pca)
    my_xml_dict = extract_meta_from_dtxml(file_name_dtxml)

    # Dictionary keys assignment. The order of the assigment will sort the xlsx file columns 
    my_dict['scan date']                           = datetime_start
    my_dict['Operator']                            = my_xml_dict['Operator']
    my_dict['Researcher']                          = my_xml_dict['Researcher']
    my_dict['NMNH PI']                             = my_xml_dict['NMNH PI']
    my_dict['Department']                          = my_xml_dict['Department']
    my_dict['CT project #']                        = my_xml_dict['Project Number']
    my_dict['specimen ID or USNM#']                = my_xml_dict['Sample ID']
    my_dict['Species name']                        = my_xml_dict['Sample Name']
    my_dict['stain']                               = my_xml_dict['Description']
    my_dict['Sample type']                         = my_xml_dict['Sample Type']
    my_dict['folder name']                         = p.stem
    # my_dict['Folder']                              = my_pcr_dict['ImageData']['pca_file']
    my_dict['timing (ms)']                         = my_pca_dict['Detector']['timingval']
    my_dict['frame avg']                           = my_pca_dict['Detector']['avg']
    my_dict['skip']                                = my_pca_dict['CT']['skipacc']
    my_dict['binning']                             = my_pca_dict['Detector']['binning']
    my_dict['sensitivity']                         = my_pca_dict['Detector']['cameragain']
    my_dict['# images']                            = my_pca_dict['CT']['numberimages']
    my_dict['total scanning time (hrs)']           = str(scan_time)
    my_dict['voltage (kV)']                        = my_pca_dict['Xray']['voltage']
    my_dict['current (uA)']                        = my_pca_dict['Xray']['current']
    my_dict['act. power (W)']                      = "Manual Entry"
    my_dict['magnification']                       = my_pca_dict['Geometry']['magnification']
    my_dict['voxel size (um)']                     = float(my_pca_dict['Geometry']['voxelsizex']) * 1000.0
    my_dict['tube type']                           = my_pca_dict['Xray']['name']
    my_dict['mode']                                = my_pca_dict['Xray']['mode']
    my_dict['multiscan (# of scans)']              = my_pca_dict['Multiscan']['active']
    my_dict['filter']                              = "Manual Entry"
    my_dict['target']                              = "Manual Entry"
    my_dict['collimator']                          = "Manual Entry"

    # Specific to Freya
    my_dict['Trawl number or collection event ID'] = "Manual Entry"
    my_dict['same specimen as Scan x']             = "Manual Entry"
    my_dict['specimen pixel width']                = "Manual Entry"

    # additional meta data
    my_dict['general system name']                 = my_pca_dict['General']['systemname']

    df = pd.DataFrame(data={**my_dict}, index=[0])
    xlsx = pathlib.Path(file_name_xlsx)
    if xlsx.is_file():
        wb = load_workbook(filename = file_name_xlsx)
        ws = wb["Sheet1"]
        for r in dataframe_to_rows(df, index=False, header=False):  #No index and don't append the column headers
            ws.append(r)
        wb.save(file_name_xlsx)
        # df.to_excel(file_name_xlsx, header=False, index=False)
        print('Append to existing meta-data excel file at: %s' % file_name_xlsx)
    else:
        df.to_excel(file_name_xlsx, header=True, index=False)
        print('Create a new meta-data excel file at: %s' % file_name_xlsx)

if __name__ == "__main__":
   main(sys.argv)