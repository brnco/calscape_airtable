# calscape_airtable
resources for converting CalScape data for use in Airtable

# Installation

1. fork / clone the repo

2. set up a virtual environment in the repo directory

   `python -m venv venv`

3. activate virtual env

   Windows: `venv\Scripts\activate.bat`
   
   Mac/ Linux: `source venv/bin/activate`
   
4. install dependencies from requirements.txt

   `python -m pip install -r requirements.txt`

# Use

1. generate plant list on CalScape

2. export detailed plant list to Excel

   must be "Detailed" version

   don't modify this Excel file, script expects the raw export .xls from CalScape

3. in Airtable:

   go to [this link](https://airtable.com/shrtRHFS0ksACljSv)

   in the upper right-hand corner, select `Use This Data`

   choose a workspace from your Airtable account, choose a base

   select `Copy this data into a new table, which you can edit`

   Keep the field names and the values in the `null` record exactly the same

4. on your machine:

   open the `config` file in your favorite text editor
   
   enter the base key (starts with `app` in the airtable URL)
   
   your api key (airtable.com -> Account -> API)
   
   the name of the table that you set up (default is `CalScape Edited Export Template`)
   
   save the config file

5. run the script 

   `calscape_airtable.py --calscape_export /path/to/export.xls`
