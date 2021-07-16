#!/usr/bin/env python

import os

import logging
import sys
import argparse
import subprocess
import json

import pathlib

logger = logging.getLogger(__name__)

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


    kicad_python_program = os.path.join(read_install_location('KiCad'), 'bin', 'python.exe')
    blender_program = os.path.join(read_install_location('Blender'), 'blender.exe')
else:
    kicad_python_program = 'python'
    blender_program = 'blender'


def mkdir_p(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path


def _main(argv=sys.argv):
    logging.basicConfig(stream=sys.stderr, level=logging.DEBUG,
                        format='%(asctime)s - %(threadName)s - %(name)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(prog=argv[0], description='KicadBlender',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-v', '--verbose', dest='verbose_count', action='count', default=0,
                        help="increases log verbosity for each occurrence.")
    parser.add_argument('-q', '--quality', default=100, type=int,
                        help="texture quality")
    parser.add_argument('-o', '--output', default=None,
                        help="output file")
    parser.add_argument('input',
                        help="input kicad file")
    parser.add_argument('output',
                        help="output directory")

    # Parse
    args, unknown_args = parser.parse_known_args(argv[1:])

    # Set logging level
    logging.getLogger().setLevel(max(3 - args.verbose_count, 0) * 10)

    verbose_args = 'v' * args.verbose_count
    if len(verbose_args):
        verbose_args = ['-' + verbose_args]
    else:
        verbose_args = []

    exec_env = os.environ.copy()
    for v in ['VIRTUAL_ENV', 'PYTHONPATH', 'PYTHONUNBUFFERED']:
        exec_env.pop(v, None)

    tmp_path = mkdir_p(os.path.join(args.output, "tmp"))
    textures_path = mkdir_p(os.path.join(args.output, "textures"))
    script_path = pathlib.Path(__file__).parent.resolve()
    wrl_file = os.path.join(args.output, os.path.basename(os.path.splitext(args.input)[0] + '.wrl'))
    blender_file = os.path.join(args.output, os.path.basename(os.path.splitext(args.input)[0] + '.blend'))

    logger.info("Export boards SVGs")
    # Call kicad script with kicad python executable
    sub_args = [kicad_python_program, os.path.join(script_path, "kicad.py")]
    sub_args += verbose_args
    sub_args += [args.input, tmp_path]
    logger.debug("Call " + " ".join(sub_args))
    subprocess.call(sub_args, env=exec_env)

    logger.debug("Opening SVG data output")
    with open(os.path.join(tmp_path, "data.json")) as json_file:
        data = json.load(json_file)

    logger.debug("Data: %s" % data)
    logger.info("SVGs to PNGs")
    svg_files = pathlib.Path(tmp_path).glob('*.svg')
    for path in svg_files:
        svg_file = str(path)
        png_file = os.path.join(textures_path, path.with_suffix(".png").name)

        # Call inkscape script with current python env
        sub_args = [sys.executable, os.path.join(script_path, "inkscape.py")]
        sub_args += verbose_args
        sub_args += [svg_file, "-o", png_file]
        sub_args += ["-s", str(args.quality)]
        sub_args += ["-u", data['units']]
        sub_args += ["%f:%f:%f:%f" % (data['x'], data['y'], data['width'], data['height'])]
        logger.debug("Call " + " ".join(sub_args))
        subprocess.call(sub_args, env=exec_env)

    logger.info("Create Blender file")
    # Call blender script with blender executable
    sub_args = [blender_program, "--background", "--python", os.path.join(script_path, "blender.py"), '--']
    sub_args += verbose_args
    sub_args += [os.path.join(script_path, "Template.blend"), wrl_file, textures_path, blender_file]
    logger.debug("Call " + " ".join(sub_args))
    subprocess.call(sub_args, env=exec_env)

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
