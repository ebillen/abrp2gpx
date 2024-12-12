# abrp2gpx
Convert tracks recorded by ABRP from xlsx to gpx

## Purpose
ABRP (https://abetterrouteplanner.com/) is able to track and save the driven route.
The saved data can be exported in xlsx-format (MS Excel).
However, I'd like to use in in gpx format.

This little tool converts the ABRP-xlsx-file to a gpx track file.

## Requirements
- python 3
- openpyxl package

The script is tested with python 3.6 in an Linux environment. However, it should run in
Wind***ws environments also, but this is not tested.

## Installation
### Linux environments
Download and install the script. If needed, make it an executable using the command "chmod +x abrp2gpx.py".
### Windows environments
Download the script.
Install python3 from MS Store.
Open a command prompt and install hte openpyxl package "openpyxl" using the command "pip install openpyxl"
In the command prompt, change to the directory where you downloaded the script and run "python3 abrp2gpx.py -h" to test the installation.

## Usage
./abrp2gpx.py -i "input_file.xlsx"

Call
./abrp2gpx.py -h
for more command line options.

