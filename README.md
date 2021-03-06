# CLI interface to consume Geocoding API

This repository contains the source code to read the address from Excel sheet, consume Geocoding API and write the latitude and longitude and formatted address fields to Excel sheet

## Requirements

To install and run these example you need:
- Python 3.3+
- git (only to clone this repository)
- Google Geocoding API key
- Excel sheet with an address field (One or more rows allowed)
- Following packages required
    - openpyxl
    - requests

## Installation

The commands below set everything up to run the examples:
```
$ git clone https://github.com/nsayed123/signzy-geocoding-project.git
$ cd signzy-geocoding-project
```

Install the following packages

- pip3 install openpyxl
- pip3 install requests

or

- pip3 install -r requirements.txt

## Get Google Geocoding API key

Instruction to get the API key can be found at
https://developers.google.com/maps/documentation/geocoding/get-api-key
```
Find the word "api_key" in the geocode.py file, update it with the generated API key
```
## Run

Make sure you are in the right directory "signzy-geocoding-project"
```
python3 geocode.py <Excel sheet>

Ex: python3 geocode.py region.xlsx
```
Note: the first row in the Excel sheet will have the field name as "Region". Example file is provided "region.xlsx"




