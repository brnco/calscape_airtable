#!/usr/bin/env python
'''
resources for converting CalScape data for use in Airtable
'''
from pathlib import Path
import olefile
import argparse
import configparser
import re
import traceback
import xlrd
from airtable import Airtable
import pandas as pd

class dotdict(dict):
    '''
    dot notation for dictionary/ object
    '''
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__

def upload_airtable_record(our_record, airtable):
    '''
    uploads record to airtable
    if record already exists, update() is used
    '''
    results = airtable.search("Current Botanical Name", our_record["Current Botanical Name"])
    if results:
        airtable.update(results[0]["id"],our_record, typecast=True)
    else:
        airtable.insert(our_record,typecast=True)

def parse_soil_fields(airtable_record):
    '''
    separates single soil field into 'tolerates' and 'prefers'
    '''
    soil_prefers = None
    soil_tolerates = None
    og_soil = airtable_record["Soil"]
    tolerates = ["tolerates", "tolerant of","accepts"]
    regx_tolerates = []
    for tol in tolerates:
        regx_tolerates.append("(" + tol + r".+\Z)")
    prog = re.compile("|".join(regx_tolerates), re.IGNORECASE)
    result = prog.search(og_soil)
    if result:
        soil_tolerates = result.group()
        tquery = result.group(1)
    prefers = ["prefers", "does best in", "does best with"]
    regx_prefers = []
    for pref in prefers:
        regx_prefers.append("(" + pref + r".+\Z)")
    prog = re.compile("|".join(regx_prefers), re.IGNORECASE)
    result = prog.search(og_soil)
    if result:
        soil_prefers = result.group()
        pquery = result.group(1)    
    if soil_prefers and soil_tolerates:
        airtable_record['Soil - Tolerates'] = soil_tolerates.replace(soil_prefers,"")
        airtable_record['Soil - Prefers'] = soil_prefers.replace(soil_tolerates,"")
    elif soil_prefers:
        airtable_record["Soil - Prefers"] = soil_prefers
    elif soil_tolerates:
        airtable_record["Soil - Tolerates"] = soil_tolerates
    if not soil_prefers:
        airtable_record.pop("Soil - Prefers", None)
    if not soil_tolerates:
        airtable_record.pop("Soil - Tolerates", None)
    return airtable_record

def parse_width_feet(w_dimensions, airtable_record):
    '''
    parses width field for imperial measurements
    '''
    w_range_ft = w_dimensions[0].split(" - ")
    airtable_record["Mature Width - min (ft)"] = float(w_range_ft[0].strip())
    if len(w_range_ft) > 1:
        airtable_record["Mature Width - max (ft)"] = float(w_range_ft[1].strip())
    else:
        airtable_record["Mature Width - max (ft)"] = float(w_range_ft[0].strip())
        airtable_record.pop("Mature Width - min (ft)",None)
    return airtable_record

def parse_width_meters(w_dimensions, airtable_record):
    '''
    parses width field for metric measurements
    '''
    w_range_m = w_dimensions[1].split(" - ")
    airtable_record["Mature Width - min (m)"] = float("".join(w_range_m[0].replace(" m)","").split()))
    if len(w_range_m) > 1:
        airtable_record["Mature Width - max (m)"] = float("".join(w_range_m[1].replace(" m)","").split()))
    else:
        airtable_record["Mature Width - max (m)"] = float("".join(w_range_m[0].replace(" m)","").split()))
        airtable_record.pop("Mature Width - min (m)", None)
    return airtable_record

def parse_height_feet(h_dimensions, airtable_record):
    '''
    parses height info for imperial measurements
    '''
    h_range_ft = h_dimensions[0].split(" - ")
    airtable_record["Mature Height - min (ft)"] = float(h_range_ft[0].strip())
    if len(h_range_ft) > 1:
        airtable_record["Mature Height - max (ft)"] = float(h_range_ft[1].strip())
    else:
        airtable_record["Mature Height - max (ft)"] = float(h_range_ft[0].strip())
        airtable_record.pop("Mature Height - min (ft)", None)
    return airtable_record

def parse_height_meters(h_dimensions, airtable_record):
    '''
    parses height info for metric measurements
    '''
    h_range_m = h_dimensions[1].split(" - ")
    airtable_record["Mature Height - min (m)"] = float(h_range_m[0].replace(" m)","").strip())
    if len(h_range_m) > 1:
        airtable_record["Mature Height - max (m)"] = float(h_range_m[1].replace(" m)","").strip())
    else:
        airtable_record["Mature Height - max (m)"] = float(h_range_m[0].replace(" m)","").strip())
        airtable_record.pop("Mature Height - min (m)", None)
    return airtable_record

def parse_dimensions_fields(airtable_record):
    '''
    takes the height and width fields and makes them useful
    '''
    try:
        og_height = airtable_record['Height']
        h_dimensions = og_height.split("ft(")
        if len(h_dimensions) < 2:
            h_dimensions = og_height.split("ft (")
        airtable_record = parse_height_feet(h_dimensions, airtable_record)
        airtable_record = parse_height_meters(h_dimensions, airtable_record)
    except:
        print("there was an issue parsing the height info for: %s",airtable_record["Common Name"])
        print(traceback.format_exc())
        airtable_record.pop("Mature Height - min (ft)",None)
        airtable_record.pop("Mature Height - min (m)",None)
        airtable_record.pop("Mature Height - max (ft)",None)
        airtable_record.pop("Mature Height - max (m)",None)
    try:
        og_width = airtable_record["Width"]
        w_dimensions = og_width.split("ft(")
        if len(w_dimensions) < 2:
            w_dimensions = og_width.split("ft (")
        if len(w_dimensions) < 2:
            #assumes CalScape is using feet
            airtable_record = parse_width_feet(w_dimensions, airtable_record)
            airtable_record.pop("Mature Width - min (m)",None)
            airtable_record.pop("Mature Width - max (m)",None)
        else:
            airtable_record = parse_width_feet(w_dimensions, airtable_record)
            airtable_record = parse_width_meters(w_dimensions, airtable_record)
    except:
        print("there was an issue parsing the width info for: %s", airtable_record["Common Name"])
        print(traceback.format_exc())
        airtable_record.pop("Mature Width - min (ft)",None)
        airtable_record.pop("Mature Width - min (m)",None)
        airtable_record.pop("Mature Width - max (ft)",None)
        airtable_record.pop("Mature Width - max (m)",None)
    return airtable_record

def lint_record(airtable_record):
    '''
    handles some other cleanup
    '''
    #airtable_record["Flowers"] = airtable_record["Flowers"].split(",")
    #airtable_record["Flowering Season"] = airtable_record["Flowering Season"].split(",")
    #airtable_record["Sun"] = airtable_record["Sun"].split(",")
    #airtable_record["Drainage"] = airtable_record["Drainage"].split(",")
    airtable_record["Popularity Ranking"] = int(airtable_record["Popularity Ranking"])
    #airtable_record["Plant Type"] = airtable_record["Plant Type"].split(",")
    #airtable_record["Growth Rate"] = airtable_record["Growth Rate"].split(",")
    #airtable_record["Water Requirement"] = airtable_record["Water Requirement"].split(",")
    #airtable_record["Ease of Care"] = airtable_record["Ease of Care"].split(",")
    #airtable_record["Availability in Nurseries"] = airtable_record["Availability in Nurseries"].split(",")
    #airtable_record["Deciduous / Evergreen"] = airtable_record["Deciduous / Evergreen"].split(",")
    #airtable_record["Common uses"] = airtable_record["Common uses"].split(",")
    return airtable_record

def get_header_columns(workbook, blank_airtable_record):
    '''
    gets the header column and maps indexes to airtable fields
    '''
    airtable_record_column_map = blank_airtable_record.copy()
    column_names = list(workbook.columns.values)
    for col in column_names:
        airtable_record_column_map[col] = column_names.index(col)
    index_column_map = \
            {y: x for x, y in airtable_record_column_map.items()}
    return index_column_map

def parse_workbook_to_airtable_record(workbook, airtable):
    '''
    takes workbook data and, per row, creates airtable records
    '''
    airtable_record = init_airtable_record(airtable)
    index_column_map = get_header_columns(workbook, airtable_record)
    lrows = workbook.values.tolist()
    for row in lrows:
        airtable_record = init_airtable_record(airtable)
        for col in row:
            if not str(col) == 'nan':
                airtable_record[index_column_map[row.index(col)]] = str(col) 
            else:
                airtable_record.pop(index_column_map[row.index(col)],None)
        airtable_record = parse_dimensions_fields(airtable_record)
        airtable_record = parse_soil_fields(airtable_record)
        airtable_record = lint_record(airtable_record)
        print(airtable_record)
        upload_airtable_record(airtable_record,airtable)
        #input("eh")

def load_calscape_export(file_path):
    '''
    loads Excel file located at file_path
    '''
    with olefile.OleFileIO(file_path) as file:
        d = file.openstream('Workbook')
        workbook = pd.read_excel(d, engine='xlrd',header=0, skiprows=4)
    return workbook

def init_airtable_record(airtable):
    '''
    initializes empty Airtable record with field names from
    https://airtable.com/shrSW1uR5Sgs670Cn
    '''
    at_recs = airtable.get_all()
    at_rec = dotdict(at_recs[0])
    del at_rec.id
    del at_rec["createdTime"]
    for field in at_rec.fields:
        field = None
    at_rec = at_rec.pop("fields")
    return at_rec

def init_airtable_connection(**kwargs):
    '''
    initializes connection to airtable using data from CLI / config
    '''
    airtable = Airtable(kwargs.base, kwargs.table, kwargs.api_key)
    return airtable

def init_kwargs(args, config):
    '''
    creates an instance of kwargs, an object which will contain
    variables and such
    '''
    kwargs = dotdict({})
    for key, val in vars(args).items():
        kwargs[key] = val
        try:
            if not val:
                kwargs[key] = config.get("Default",key)
        except configparser.NoOptionError as e:
           continue
    return kwargs

def init_config():
    '''
    loads config info (if exists)
    '''
    config_path = Path( __file__ ).absolute().with_name("config")
    if not config_path.is_file():
        return False
    config = configparser.ConfigParser()
    config.read(config_path)
    '''
    email = config.get("Default","Airtable Email Login")  
    api_key = config.get("Default","Airtable API Key")
    base = config.get("Default","Airtable Base")
    table = config.get("Default","Airtable Table")'''
    return config

def init_args():
    '''
    initialize variables and arguments
    '''
    parser = argparse.ArgumentParser(description='Process CalScape export for Airtable')
    parser.add_argument('--email', default=None, \
            help="the Airtable login email")
    parser.add_argument('--api_key', default=None, \
            help="the Airtable API key")
    parser.add_argument('--base', default=None,\
            help="the Airtable base to connect to")
    parser.add_argument('--table', default=None, \
            help="the Airtable table in the base")
    parser.add_argument('--calscape_export', type=Path, \
            help="the path to the CalScape export .xls file")
    args = parser.parse_args()
    return args

def main():
    '''
    do the thing
    '''
    try:
        args = init_args()
        config = init_config()
        kwargs = init_kwargs(args, config)
        airtable = Airtable(kwargs.base, kwargs.table, kwargs.api_key)
        calscape_export = load_calscape_export(kwargs.calscape_export)
        workbook = load_calscape_export(kwargs.calscape_export)
        parse_workbook_to_airtable_record(workbook, airtable)
    except Exception as e:
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
