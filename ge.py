
import sys
import pprint
import xmltodict
import configparser
import pandas as pd
import pathlib
from datetime import datetime

if __name__== '__main__':

    if len(sys.argv) == 1:
        print ('Must provide the path to the pca file as the argument')
        sys.exit(1)
    
    file_name_pca   = sys.argv[1]
    p = pathlib.PurePath(file_name_pca)

    # print(p.parents[0])
    # print(p.stem)
    # print(p.suffix)

    file_name_pcp   = p.with_suffix('.pcp')
    file_name_pcr   = p.with_suffix('.pcr')
    file_name_dtxml = p.with_suffix('.dtxml')
    file_name_xlsx  = p.with_suffix('.xlsx')

    my_dict = {}

    # Extract meta data from pcp file
    with open(file_name_pcp) as file:
        time_start = file.readlines()[3].split("\t")[-1].strip()
        datetime_start = datetime.strptime(time_start, '%Y-%m-%d %H:%M:%S')
        file.seek(0)  
        time_end = file.readlines()[-1].split("\t")[-1].strip()
        datetime_end = datetime.strptime(time_end, '%Y-%m-%d %H:%M:%S')

    scan_time = datetime_end - datetime_start
    my_dict['total scanning time (hrs)'] = str(scan_time)
    my_dict['scan date'] = datetime_start

    # Extract meta data from pcr file
    with open(file_name_pcr, encoding='latin-1') as f:
        config = configparser.ConfigParser()
        config.read_file(f)   

    sections = config.sections()
    my_pcr_dict = {i: {i[0]: i[1] for i in config.items(i)} for i in config.sections()}
    # for key in my_pcr_dict:
    #     print(key, my_pcr_dict[key])

    # my_dict['Folder'] = my_pcr_dict['ImageData']['pca_file']
    my_dict['Folder'] = p.stem

    # Extract meta data from pca file
    with open(file_name_pca, encoding='latin-1') as f:
        config = configparser.ConfigParser()
        config.read_file(f)   

    sections = config.sections()
    my_pca_dict = {i: {i[0]: i[1] for i in config.items(i)} for i in config.sections()}
    # for key in my_pca_dict:
    #     print(key, my_pca_dict[key])

    my_dict['timing (ms)'] = my_pca_dict['Detector']['timingval']
    my_dict['binning'] = my_pca_dict['Detector']['binning']
    my_dict['skip'] = my_pca_dict['CT']['skipacc']
    my_dict['# images'] = my_pca_dict['CT']['numberimages']
    my_dict['voltage (kV)'] = my_pca_dict['Xray']['voltage']
    my_dict['mode'] = my_pca_dict['Xray']['mode']
    my_dict['multiscan (# of scans)'] = my_pca_dict['Multiscan']['active']
    my_dict['frame avg']   = my_pca_dict['Detector']['avg']
    my_dict['general system name'] = my_pca_dict['General']['systemname']

    # Manual entries
    my_dict['collimator']   = "Manual Entry"
    my_dict['filter']   = "Manual Entry"

    # Extract meta data from dtxml file
    with open(file_name_dtxml, encoding='latin-1') as f:
    	xml_content= f.read()

    xml_dict=xmltodict.parse(xml_content)
    i = 0
    my_xml_dict = {}
    while i < len(xml_dict['project']['additional_project_info']['property']):
        my_xml_dict[(xml_dict['project']['additional_project_info']['property'][i]['@name'])] =  xml_dict['project']['additional_project_info']['property'][i]['@value']
        i = i + 1

    # pprint.pprint(my_xml_dict)
    my_dict['stain'] = my_xml_dict['Description']
    my_dict['Researcher'] = my_xml_dict['Researcher']
    my_dict['CT project #'] = my_xml_dict['Project Number']
    my_dict['Species name'] = my_xml_dict['Sample Name']
    my_dict['Species name'] = my_xml_dict['Sample Name']
    my_dict['specimen ID or USNM#'] = my_xml_dict['Sample ID']

    df = pd.DataFrame(data={**my_dict}, index=[0])
    df.to_excel(file_name_xlsx)
    print('Converted excel is at: /Users/decarlo/conda/sandbox/nocturn/data/tt.xlsx')