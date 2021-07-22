#!/usr/bin/env python

import os
import sys
# cx_freeze issue like this one https://github.com/pyinstaller/pyinstaller/issues/3795
if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.SetDllDirectoryA(None)

import pcbnew
import logging
import argparse
import json
import subprocess

logger = logging.getLogger(__name__)


def call_program(*args, **kwargs):
    exec_env = os.environ.copy()
    for v in ['VIRTUAL_ENV', 'PYTHONPATH', 'PYTHONUNBUFFERED']:
        exec_env.pop(v, None)
    kwargs = kwargs.copy()
    kwargs.update(env=exec_env)
    ret = subprocess.call(*args, **kwargs)
    if ret != 0:
        raise Exception("Error occurs in a subprogram")


def set_default_settings(board):
    plot_controller = pcbnew.PLOT_CONTROLLER(board)
    plot_options = plot_controller.GetPlotOptions()

    plot_options.SetPlotFrameRef(False)
    plot_options.SetLineWidth(pcbnew.FromMM(0.35))
    plot_options.SetScale(1)
    plot_options.SetUseAuxOrigin(True)
    plot_options.SetMirror(False)
    plot_options.SetExcludeEdgeLayer(False)
    plot_controller.SetColorMode(True)


def plot(board, layer, file, name, output_directory):
    plot_controller = pcbnew.PLOT_CONTROLLER(board)
    plot_options = plot_controller.GetPlotOptions()
    plot_options.SetOutputDirectory(output_directory)
    plot_controller.SetLayer(layer)
    plot_controller.OpenPlotfile(file, pcbnew.PLOT_FORMAT_SVG, name)
    output_filename = plot_controller.GetPlotFileName()
    plot_controller.PlotLayer()
    plot_controller.ClosePlot()
    return output_filename


def normalize(point):
    return [point[0] / pcbnew.IU_PER_MM, point[1] / pcbnew.IU_PER_MM]


def parse_poly_set(self, polygon_set):
    result = []
    for polygon_index in range(polygon_set.OutlineCount()):
        outline = polygon_set.Outline(polygon_index)
        if not hasattr(outline, "PointCount"):
            self.logger.warn("No PointCount method on outline object. "
                             "Unpatched kicad version?")
            return result
        parsed_outline = []
        for point_index in range(outline.PointCount()):
            point = outline.CPoint(point_index)
            parsed_outline.append(self.normalize([point.x, point.y]))
        result.append(parsed_outline)

    return result


def parse_shape(d):
    # type: (pcbnew.PCB_SHAPE) -> dict or None
    shape = {
        pcbnew.S_SEGMENT: "segment",
        pcbnew.S_CIRCLE: "circle",
        pcbnew.S_ARC: "arc",
        pcbnew.S_POLYGON: "polygon",
        pcbnew.S_CURVE: "curve",
        pcbnew.S_RECT: "rect",
    }.get(d.GetShape(), "")
    if shape == "":
        logger.info("Unsupported shape %s, skipping", d.GetShape())
        return None
    start = normalize(d.GetStart())
    end = normalize(d.GetEnd())
    if shape in ["segment", "rect"]:
        return {
            "type": shape,
            "start": start,
            "end": end,
            "width": d.GetWidth() * 1e-6
        }
    if shape == "circle":
        return {
            "type": shape,
            "start": start,
            "radius": d.GetRadius() * 1e-6,
            "width": d.GetWidth() * 1e-6
        }
    if shape == "arc":
        a1 = round(d.GetArcAngleStart() * 0.1, 2)
        a2 = round((d.GetArcAngleStart() + d.GetAngle()) * 0.1, 2)
        if d.GetAngle() < 0:
            (a1, a2) = (a2, a1)
        return {
            "type": shape,
            "start": start,
            "radius": d.GetRadius() * 1e-6,
            "startangle": a1,
            "endangle": a2,
            "width": d.GetWidth() * 1e-6
        }
    if shape == "polygon":
        if hasattr(d, "GetPolyShape"):
            polygons = parse_poly_set(d.GetPolyShape())
        else:
            logger.info("Polygons not supported for KiCad 4, skipping")
            return None
        angle = 0
        if hasattr(d, 'GetParentModule'):
            parent_footprint = d.GetParentModule()
        else:
            parent_footprint = d.GetParentFootprint()
        if parent_footprint is not None:
            angle = parent_footprint.GetOrientation() * 0.1,
        return {
            "type": shape,
            "pos": start,
            "angle": angle,
            "polygons": polygons
        }
    if shape == "curve":
        return {
            "type": shape,
            "start": start,
            "cpa": normalize(d.GetBezControl1()),
            "cpb": normalize(d.GetBezControl2()),
            "end": end,
            "width": d.GetWidth() * 1e-6
        }


def parse_drawing(d):
    if d.GetClass() in ["DRAWSEGMENT", "MGRAPHIC", "PCB_SHAPE"]:
        return parse_shape(d)
    else:
        return None


def parse_edges(pcb):
    edges = []
    drawings = list(pcb.GetDrawings())
    bbox = None
    for d in drawings:
        if d.GetLayer() == pcbnew.Edge_Cuts:
            parsed_drawing = parse_drawing(d)
            if parsed_drawing:
                edges.append(parsed_drawing)
                if bbox is None:
                    bbox = d.GetBoundingBox()
                else:
                    bbox.Merge(d.GetBoundingBox())
    if bbox:
        bbox.Normalize()
    return edges, bbox


def get_bounding_box(pcb):
    edges, _ = parse_edges(pcb)
    x_min = y_min = float('+inf')
    x_max = y_max = float('-inf')
    for edge in edges:
        if 'start' in edge and 'end' in edge:
            xs, ys = zip(edge['start'], edge['end'])
            for x in xs:
                x_min = min(x_min, x)
                x_max = max(x_max, x)
            for y in ys:
                y_min = min(y_min, y)
                y_max = max(y_max, y)
    return x_min, y_min, x_max, y_max


def export_layers(board, output_directory):
    set_default_settings(board)

    for layer in board.GetEnabledLayers().UIOrder():
        file = name = board.GetLayerName(layer)
        logger.debug('plotting layer {} ({}) to SVG'.format(name, layer))
        output_filename = plot(board, layer, file, name, output_directory)
        logger.info('Layer %s SVG: %s' % (name, output_filename))


def export_vrml(board_file, vrml_file, origin):
    if os.name == 'nt':
        kicad2vrml = os.path.join(os.path.dirname(sys.executable), "kicad2vrml.exe")
    else:
        kicad2vrml = "kicad2vrml"
    sub_args = [kicad2vrml]
    sub_args += [board_file]
    sub_args += ["-f", "-o", vrml_file]
    sub_args += ["--user-origin", "%fx%f" % (origin[0], origin[1])]
    logger.debug("Call " + " ".join(sub_args))
    call_program(sub_args)


def _main(argv=sys.argv):
    logging.basicConfig(stream=sys.stderr, level=logging.DEBUG,
                        format='%(asctime)s - %(threadName)s - %(name)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(prog=argv[0], description='KicadBlender',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-v', '--verbose', dest='verbose_count', action='count', default=0,
                        help="increases log verbosity for each occurrence.")
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

    logger.debug("Loading board")
    board = pcbnew.LoadBoard(args.input)
    if not os.path.exists(args.output):
        os.makedirs(args.output)

    logger.info("Export Layers to SVG")
    export_layers(board, args.output)

    x, y = board.GetGridOrigin()
    x, y = (x / pcbnew.IU_PER_MM, y / pcbnew.IU_PER_MM)
    vrml_file = os.path.join(args.output, os.path.basename(os.path.splitext(args.input)[0] + '.wrl'))
    logger.info("Export VRML")
    export_vrml(args.input, vrml_file, (x, y))

    with open(os.path.join(args.output, "data.json"), 'wb') as fp:
        x1, y1, x2, y2 = get_bounding_box(board)
        values = {
            'x': x1,
            "y": y1,
            'width': x2 - x1,
            'height': y2 - y1,
            'thickness': board.GetDesignSettings().GetBoardThickness() / pcbnew.IU_PER_MM,
            'units': 'mm',
            'vrml': vrml_file
        }
        json.dump(values, fp)


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
