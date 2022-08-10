# calscape_airtable
resources for converting CalScape data for use in Airtable

# Installation

fork / clone per usual

run the virtual env

install from requirements.txt

get an API key from Airtable

# Use

generate plant list on CalScape

export plant list to Excel

don't modify this Excel file

Copy Airtable template from [this url](https://airtable.com/shrtRHFS0ksACljSv) to your own Airtable account

Keep the field names and the values in the `null` record exactly the same

open the `config` file in your favorite text editor and enter the base key (starts with `app`), your api key, and the name of the table that you set up (i.e. the name of the table that the above template is copied to). save the config file

`calscape_airtable.py --calscape_export /path/to/export.xls`
