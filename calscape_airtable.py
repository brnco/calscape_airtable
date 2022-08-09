#!/usr/bin/env python
'''
resources for converting CalScape data for use in Airtable
'''
from pathlib import Path
import argparse
import configparser
import re

class dotdict(dict):
    '''
    dot notation for dictionary/ object
    '''
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__

def init_kwargs(args, config):
    '''
    creates an instance of kwargs, an object which will contain
    variables and such
    '''
    kwargs = dotdict({})
    for key, val in vars(args).items():
        kwargs[key] = val
        if not val:
            kwargs[key] = config.get("Default",key)
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
    args = parser.parse_args()
    return args

def main():
    '''
    do the thing
    '''
    args = init_args()
    config = init_config()
    kwargs = init_kwargs(args, config)
    print(kwargs)

if __name__ == "__main__":
    main()
