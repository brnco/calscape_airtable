# calscape_airtable
resources for converting CalScape data for use in Airtable

# Installation

fork / clone per usual

run the virtual env

install from requirements.txt

get an API key from Airtable

# Use

1. generate plant list on CalScape

   export plant list to Excel

   don't modify this Excel file

2. in Airtable:

   go to [this link](https://airtable.com/shrtRHFS0ksACljSv)

   in the upper right-hand corner, select `Use This Data`

   choose a workspace from your account, choose a base

   select "Copy this data into a new table, which you can edit"

   Keep the field names and the values in the `null` record exactly the same

3. on your machine:

   open the `config` file in your favorite text editor
   
   enter the base key (starts with `app` in the airtable URL)
   
   your api key (airtable.com -> Account -> API)
   
   the name of the table that you set up (default is `CalScape Edited Export Template`)
   
   save the config file

4. run the script 

   `calscape_airtable.py --calscape_export /path/to/export.xls`
