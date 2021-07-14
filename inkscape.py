#!/usr/bin/env python

import xml.dom.minidom as minidom
import re
import sys
import subprocess
import argparse
import logging
import locale
import os

from enum import Enum

logger = logging.getLogger(__name__)

number = r'([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)'
number_unit = number + '(\w+)?'


class LengthUnit(Enum):
    mm = ('mm', 3.779528)
    cm = ('cm', 37.79528)


str_to_enum_unit = {a.value[0]: a.value for a in LengthUnit}

if os.name == 'nt':
    import winreg


    def read_install_location(software_name):
        # Need to traverse the two registry
        sub_key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall']

        for i in sub_key:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, i, 0, winreg.KEY_READ)
            for j in range(0, winreg.QueryInfoKey(key)[0] - 1):
                try:
                    key_name = winreg.EnumKey(key, j)
                    key_path = i + '\\' + key_name
                    each_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
                    displayName, _ = winreg.QueryValueEx(each_key, 'DisplayName')
                    if software_name.lower() in displayName.lower():
                        installLocation, _ = winreg.QueryValueEx(each_key, 'InstallLocation')
                        return installLocation
                except WindowsError:
                    pass
        raise Exception("Install location not found for %s" % software_name)


    ink_program = os.path.join(read_install_location('Inkscape'), 'bin', 'inkscape.com')
else:
    ink_program = 'inkscape'

# Use current locale (used by inkscape)
curr = locale.getdefaultlocale()
locale.setlocale(locale.LC_ALL, curr[0])


def number_format(number):
    str = '{0:n}'.format(number)
    return re.sub(r"\s+", '', str)  # Remove space separators


class Transform:
    def __init__(self, **entries):
        self.__dict__.update(entries)


def svg_to_png(ifile, ofile, options):
    doc = minidom.parse(ifile)
    svg_elem = doc.getElementsByTagName('svg')[0]
    height = svg_elem.getAttribute('height')
    match = re.search(number_unit, height, re.IGNORECASE)
    if not match:
        raise Exception("Invalid height attribute")
    user_to_dpi_r = str_to_enum_unit[options.unit][1]
    ink_x1 = options.x * user_to_dpi_r
    if os.name != 'nt':
        height_value = float(match.group(1))
        height_unit = str_to_enum_unit[match.group(2)]
        svg_to_dpi_r = height_unit[1]
        ink_y1 = height_value * svg_to_dpi_r - (options.y + options.height) * user_to_dpi_r
    else:
        ink_y1 = options.y * user_to_dpi_r
    ink_width = options.width * user_to_dpi_r
    ink_height = options.height * user_to_dpi_r
    ink_x2 = ink_x1 + ink_width
    ink_y2 = ink_y1 + ink_height
    ink_scale = options.scale / user_to_dpi_r * 96
    args = [ink_program]
    args += ['--export-dpi=%s' % number_format(ink_scale)]
    args += ['--export-area=%s:%s:%s:%s' % tuple([number_format(a) for a in (ink_x1, ink_y1, ink_x2, ink_y2)])]
    args += ['--export-background=white']
    args += ['-o' if os.name == 'nt' else '-e', ofile, ifile]
    logger.debug("Call " + " ".join(args))
    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def coord_type(strings):
    return tuple(map(float, strings.split(":")))


def _main(argv=sys.argv):
    logging.basicConfig(stream=sys.stderr, level=logging.DEBUG,
                        format='%(asctime)s - %(threadName)s - %(name)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(prog=argv[0], description='KicadBlender',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-v', '--verbose', dest='verbose_count', action='count', default=0,
                        help="increases log verbosity for each occurrence.")
    parser.add_argument('-o', '--output', default=None,
                        help="output file")
    parser.add_argument('-s', '--scale', default=100.0, type=float,
                        help="scale to apply (pixel by user unit)")
    parser.add_argument('-u', '--unit', default="cm",
                        help="user unit (taken from SVG if not defined)")
    parser.add_argument('input',
                        help="input file")
    parser.add_argument('area', type=coord_type,
                        help="area to export (user unit) format: x:y:w:h")

    # Parse
    args, unknown_args = parser.parse_known_args(argv[1:])

    # Set logging level
    logging.getLogger().setLevel(max(3 - args.verbose_count, 0) * 10)

    if args.output is None:
        args.output = os.path.splitext(args.input)[0] + '.png'

    options = Transform(**{
        'x': args.area[0],
        'y': args.area[1],
        'width': args.area[2],
        'height': args.area[3],
        'scale': args.scale,
        'unit': args.unit,
    })

    logger.debug('Convert SVG %s to PNG %s' % (args.input, args.output))
    svg_to_png(args.input, args.output, options)


def main():
    try:
        sys.exit(_main(sys.argv))
    except Exception as e:
        logger.exception(e)
        sys.exit(-1)
    finally:
        logging.shutdown()


if __name__ == "__main__":
    main()
