# calscape_airtable
resources for converting CalScape data for use in Airtable

# Installation

fork and clone per usual

run the virtual env

get an API key from Airtable

# Use

generate plant list on CalScape

export plant list to Excel or CSV

`calscape_airtable.py --calscape_export /path/to/export.csv --airtable_url airtable.com/abcd/1234 --user email@example.com --api_key abcd1234`

The above uploads the data directly and assumes you're using the template located at [this link](https://airtable.com/shrMir3FeAZCPbgD6)
