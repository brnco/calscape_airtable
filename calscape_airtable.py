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

def process_dimensions_fields(airtable_record):
    '''
    takes the height and width fields and makes them useful
    '''
    try:
        og_height = airtable_record['Height']
        h_dimensions = og_height.split("ft(")
        if len(h_dimensions) < 2:
            h_dimensions = og_height.split("ft (")
        h_range_ft = h_dimensions[0].split(" - ")
        h_range_m = h_dimensions[1].split(" - ")
        airtable_record["Mature Height - min (ft)"] = float(h_range_ft[0].strip())
        airtable_record["Mature Height - min (m)"] = float(h_range_m[0].strip())
        airtable_record["Mature Height - max (ft)"] = float(h_range_ft[1].strip())
        airtable_record["Mature Height - max (m)"] = float(h_range_m[1].replace(" m)","").strip())
    except:
        print("there was an issue parsing the height info for: %s",airtable_record["Common Name"])
        print(traceback.format_exc())
        del airtable_record["Mature Height - min (ft)"]
        del airtable_record["Mature Height - min (m)"]
        del airtable_record["Mature Height - max (ft)"]
        del airtable_record["Mature Height - max (m)"]
    try:
        og_width = airtable_record["Width"]
        w_dimensions = og_width.split("ft(")
        if len(w_dimensions) < 2:
            w_dimensions = og_width.split("ft (")
        w_range_ft = w_dimensions[0].split(" - ")
        w_range_m = w_dimensions[1].split(" - ")
        airtable_record["Mature Width - min (ft)"] = float(w_range_ft[0].strip())
        airtable_record["Mature Width - min (m)"] = float(w_range_m[0].strip())
        airtable_record["Mature Width - max (ft)"] = float(w_range_ft[1].strip())
        airtable_record["Mature Width - max (m)"] = float(w_range_m[1].replace(" m)","").strip())
    except:
        print("there was an issue parsing the width info for: %s", airtable_record["Common Name"])
        print(traceback.format_exc())
        del airtable_record["Mature Width - min (ft)"]
        del airtable_record["Mature Width - min (m)"]
        del airtable_record["Mature Width - max (ft)"]
        del airtable_record["Mature Width - max (m)"]
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
        for col in row:
            airtable_record[index_column_map[row.index(col)]] = col 
        airtable_record = process_dimensions_fields(airtable_record)
        print(airtable_record)
        input("eh")
    return airtable_record

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
        something = parse_workbook_to_airtable_record(workbook, airtable)
    except Exception as e:
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
