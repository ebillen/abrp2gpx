#!/usr/bin/env python3

import argparse
import datetime
import logging
import logging.handlers
import openpyxl
from time import mktime, strptime
from xml.etree.ElementTree import Element, SubElement, ElementTree

__version__ = '0.1 (EB)'

# structure to hold the data from the input file:
abrp = {}
# sheetname in the ABRP output:
abrp['sheetName'] = 'ABRP Activity'
# we will store the waypoints from the input file here:
abrp['waypoints'] = []



def main(args):
    logger.debug('main()')
    logger.debug('  input file:  {}'.format(args.input))
    logger.debug('  output file: {}'.format(args.output))

    # Try to load the excel file:
    try:
        input_excel = openpyxl.load_workbook(args.input)
    except FileNotFoundError:
        logger.error('Failed to open "{}" for reading - '
                     'aborting!'.format(args.input))
        exit(1)

    # Try to get the (only) sheet from the excel:
    try:
        sheet = input_excel[abrp['sheetName']]
    except KeyError:
        logger.error('No sheet "{}" in "{}" - is this a ABRP export?'.format(
            abrp['sheetName'],
            args.input))
        exit(1)

    abrp['title'] = sheet['A1'].value
    abrp['km_start'] = sheet['K5'].value
    abrp['km_end'] = sheet['L5'].value
    logger.info('Input file: "{}"'.format(args.input))
    logger.info('Title: "{}"'.format(abrp['title']))
    logger.info('Odometer: {} - {}'.format(abrp['km_start'], abrp['km_end']))

    # Read the trip data from the xlsx-file:

    numWaypoints = 0
    for row in sheet.iter_rows(min_row=9):
        timestampstring = row[0].value
        # timestampstring is in format "1.11.2024, 14:28:39"
        # TODO: the time format probably depends on user's language settings
        # in ABRP...
        timestamp = mktime(strptime(timestampstring, '%d.%m.%Y, %H:%M:%S'))
        logger.debug('  timestamp is {}'.format(timestampstring))
        logger.debug('    => {}'.format(timestamp))
        lat = row[1].value
        lon = row[2].value
        height = row[4].value
        abrp['waypoints'].append({'lat': lat,
                                  'lon': lon,
                                  'height': height,
                                  'timestamp': timestamp})
        numWaypoints += 1

    logger.info('Read {} waypoints from input file'.format(numWaypoints))

    # Create gpx from the data:
    gpx = Element('gpx',
                  version="1.1",
                  creator='abrp2gpx.py version {}'.format(__version__),
                  xmlns="http://www.topografix.com/GPX/1/1")
    metadata = SubElement(gpx, 'metadata')
    time = SubElement(metadata, 'time')
    time.text = datetime.datetime.utcnow().isoformat() + "Z"
    description = SubElement(metadata, 'description')
    description.text = abrp['title']

    trk = SubElement(gpx, 'trk')
    trkseg = SubElement(trk, 'trkseg')

    for wp in abrp['waypoints']:
        wpt = SubElement(trkseg,
                         'trkpt',
                         lat=str(wp['lat']),
                         lon=str(wp['lon']))
        if wp.get('name'):
            name = SubElement(wpt, 'name')
            name.text = str(wp['name'])
        if wp.get('timestamp'):
            time = SubElement(wpt, 'time')
            time.text = datetime.datetime.fromtimestamp(wp['timestamp']).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"

    tree = ElementTree(gpx)
    try:
        tree.write(args.output,
                   encoding='utf-8',
                   xml_declaration=True)
    except PermissionError:
        logger.error('Failed to write to "{}" - '
                     'permission denied!'.format(args.output))
        exit(1)

    logger.info('Wrote gpx file to "{}".'.format(args.output))


if __name__ == '__main__':
    # prepare console logger:
    logger = logging.getLogger(__name__)
    logger.setLevel(level=logging.DEBUG)

    formatter = logging.Formatter('%(asctime)-25s%(levelname)-10s%(message)s')

    consoleHandler = logging.StreamHandler()
    # default verbosity: INFO
    # command line option '-d' and '-q' will change this
    consoleHandler.setLevel(logging.INFO)
    consoleHandler.setFormatter(formatter)

    logger.addHandler(consoleHandler)

    parser = argparse.ArgumentParser(description='ABRP2GPX')
    parser.add_argument(
        '-v', '--version',
        action='version',
        version='%(prog)s ' + __version__)
    parser.add_argument(
        '-d', '--debug',
        action='store_true',
        dest='debug',
        help='debug mode - print debug output to console')
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        dest='quiet',
        help='quiet mode - print only errors to console')
    parser.add_argument(
        '-i', '--input',
        action='store',
        required=True,
        type=str,
        help='Input excel file')
    parser.add_argument(
        '-o', '--output',
        action='store',
        help='Output gpx file (default: inputfile.gpx)')

    args = parser.parse_args()

    if args.debug:
        consoleHandler.setLevel(logging.DEBUG)
    if args.quiet:
        consoleHandler.setLevel(logging.ERROR)
    if not args.output:
        args.output = args.input.replace('.xlsx', '.gpx')

    main(args)
