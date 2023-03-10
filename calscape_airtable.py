#!/usr/bin/env python
'''
resources for converting CalScape data for use in Airtable
'''

import argparse
import configparser
import traceback
import re
from pprint import pprint
from pathlib import Path
import olefile
import time

from airtable import Airtable
import pandas as pd


class dotdict(dict):
    '''
    dot notation for dictionary/ object
    '''
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__

def upload_airtable_record(our_record, airtbl):
    '''
    uploads record to airtable
    if record already exists, update() is used
    '''
    results = airtbl.search("Current Botanical Name", our_record["Current Botanical Name"])
    if results:
        airtbl.update(results[0]["id"],our_record, typecast=True)
    else:
        airtbl.insert(our_record,typecast=True)

def parse_soil_fields(airtable_record):
    '''
    separates single soil field into 'tolerates' and 'prefers'
    '''
    try:
        if airtable_record["Soil"]:
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
                #tquery = result.group(1)
            prefers = ["prefers", "does best in", "does best with"]
            regx_prefers = []
            for pref in prefers:
                regx_prefers.append("(" + pref + r".+\Z)")
            prog = re.compile("|".join(regx_prefers), re.IGNORECASE)
            result = prog.search(og_soil)
            if result:
                soil_prefers = result.group()
                #pquery = result.group(1)
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
    except KeyError:
        pass
    return airtable_record

def parse_width_feet(w_dimensions, airtable_record):
    '''
    parses width field for imperial measurements
    '''
    w_range_ft = w_dimensions[0].split(" - ")
    if len(w_range_ft) < 2:
        w_range_ft = w_range_ft[0].split("-")
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
    if len(w_range_m) < 2:
        w_range_m = w_range_m[0].split("-")
    airtable_record["Mature Width - min (m)"] = \
        float("".join(w_range_m[0].replace(" m)","").split()))
    if len(w_range_m) > 1:
        airtable_record["Mature Width - max (m)"] = \
            float("".join(w_range_m[1].replace(" m)","").split()))
    else:
        airtable_record["Mature Width - max (m)"] = \
            float("".join(w_range_m[0].replace(" m)","").split()))
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

def parse_inches_centimeters(dimensions):
    '''
    converts inches / centimeters to ft/m
    '''
    processed_dimensions = []
    try:
        dimensions = dimensions[0].split(" in(")
        if len(dimensions) < 2:
            dimensions = dimensions[0].split(" in (")
        if len(dimensions) < 2:
            return dimensions
        inches = dimensions[0].split(" - ")
        if len(inches) < 2:
            inches = dimensions[0].split("-")
        if len(inches) < 2:
            feet_max = float(inches[0].strip()) / 12.0
            processed_dimensions.append(str(feet_max))
        if len(inches) == 2:
            feet_min = float(inches[0].strip()) / 12.0
            feet_max = float(inches[1].strip()) / 12.0
            processed_dimensions.append(str(feet_min) + " - " + str(feet_max))
        centimeters = dimensions[1].split(" - ")
        if len(centimeters) < 2:
            centimeters = dimensions[1].split("-")
        if len(centimeters) < 2:
            meters_max = float(centimeters[0].replace(")","").replace("cm","").strip()) / 100.0
            processed_dimensions.append(str(meters_max))
        if len(centimeters) == 2:
            meters_min = float(centimeters[0]) / 100.0
            meters_max = float(centimeters[1].replace(")","").replace("cm","").strip()) / 100.0
            processed_dimensions.append(str(meters_min) + " - " + str(meters_max))
        return processed_dimensions
    except:
        print("there was an issue converting to feet/meters")
        print(traceback.format_exc())
        return dimensions

def parse_dimensions_fields(airtable_record):
    '''
    takes the height and width fields and makes them useful
    '''
    try:
        og_height = airtable_record['Height']
    except KeyError:
        og_height = None
    try:
        h_dimensions = og_height.split("ft(")
        if len(h_dimensions) < 2 and ("in" in h_dimensions[0] or "cm" in h_dimensions[0]):
            h_dimensions = parse_inches_centimeters(h_dimensions)
        if len(h_dimensions) < 2:
            h_dimensions = og_height.split("ft (")
        airtable_record = parse_height_feet(h_dimensions, airtable_record)
        airtable_record = parse_height_meters(h_dimensions, airtable_record)
    except:
        print("there was an issue parsing the height info for: ",\
            airtable_record["Current Botanical Name"])
        #print(traceback.format_exc())
        airtable_record.pop("Mature Height - min (ft)",None)
        airtable_record.pop("Mature Height - min (m)",None)
        airtable_record.pop("Mature Height - max (ft)",None)
        airtable_record.pop("Mature Height - max (m)",None)
    try:
        og_width = airtable_record["Width"]
    except KeyError:
        og_width = None
    try:
        w_dimensions = og_width.split("ft(")
        if len(w_dimensions) < 2 and ("in" in w_dimensions[0] or "cm" in w_dimensions[0]):
            w_dimensions = parse_inches_centimeters(w_dimensions)
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
        print("there was an issue parsing the width info for: ", \
            airtable_record["Current Botanical Name"])
        #print(traceback.format_exc())
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
    airtable_record["Popularity Ranking"] = \
        int(airtable_record["Popularity Ranking"])
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

def parse_workbook_to_airtable_record(workbook, airtbl):
    '''
    takes workbook data and, per row, creates airtable records
    '''
    blank_airtable_record = init_airtable_record(airtbl)
    index_column_map = get_header_columns(workbook, blank_airtable_record)
    lrows = workbook.values.tolist()
    for row in lrows:
        airtable_record = init_airtable_record(airtbl)
        for col in row:
            if str(col) != 'nan':
                airtable_record[index_column_map[row.index(col)]] = str(col)
            else:
                airtable_record.pop(index_column_map[row.index(col)],None)
        airtable_record = parse_dimensions_fields(airtable_record)
        airtable_record = parse_soil_fields(airtable_record)
        airtable_record = lint_record(airtable_record)
        pprint(airtable_record)
        print("uploading above info for " + airtable_record["Current Botanical Name"])
        upload_airtable_record(airtable_record,airtbl)
        time.sleep(0.25)

def load_calscape_export(file_path):
    '''
    loads Excel file located at file_path
    '''
    with olefile.OleFileIO(file_path) as file:
        stream = file.openstream('Workbook')
        workbook = pd.read_excel(stream, engine='xlrd',header=0, skiprows=4)
    return workbook

def init_airtable_record(airtbl):
    '''
    initializes empty Airtable record with field names from
    https://airtable.com/shrSW1uR5Sgs670Cn
    '''
    at_recs = airtbl.get_all()
    at_rec = dotdict(at_recs[0])
    del at_rec.id
    del at_rec["createdTime"]
    for field in at_rec['fields']:
        at_rec['fields'][field] = None
    at_rec = at_rec.pop("fields")
    return at_rec

def init_airtable_connection(kwvars):
    '''
    initializes connection to airtable using data from CLI / config
    '''
    airtbl = Airtable(kwvars.base, kwvars.table, kwvars.api_key)
    return airtbl

def init_kwvars(args, config):
    '''
    creates an instance of kwvars, an object which will contain
    variables and such
    '''
    kwvars = dotdict({})
    for key, val in vars(args).items():
        kwvars[key] = val
        try:
            if not val:
                kwvars[key] = config.get("Default",key)
        except configparser.NoOptionError:
            continue
    return kwvars

def init_config():
    '''
    loads config info (if exists)
    '''
    config_path = Path( __file__ ).absolute().with_name("config")
    if not config_path.is_file():
        return False
    config = configparser.ConfigParser()
    config.read(config_path)
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
    print("starting...")
    try:
        args = init_args()
        config = init_config()
        kwvars = init_kwvars(args, config)
        airtbl = Airtable(kwvars.base, kwvars.table, kwvars.api_key)
        workbook = load_calscape_export(kwvars.calscape_export)
        parse_workbook_to_airtable_record(workbook, airtbl)
    except Exception:
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
