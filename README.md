# CLI interface to consume Geocoding API

This repository contains the source code to read the address from Excel sheet, consume Geocoding API and write the result to Excel sheet

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
$ git clone https://github.com/nsayed123/signzy-geocoding-project.git
$ cd signzy-geocoding-project

Install the following packages
pip3 install openpyxl
pip3 install requests

or

pip3 install -r requirements.txt

## Get Google Geocoding API key

Instruction to get the API key
https://developers.google.com/maps/documentation/geocoding/get-api-key

At line 62 / find the word "api_key" in the code update the generated API key

## Run

Make sure your current directory signzy-geocoding-project
python3 geocode.py <Excel sheet>

Ex: python3 geocode.py region.xlsx

Note: the first row will have the field name as "Region"




