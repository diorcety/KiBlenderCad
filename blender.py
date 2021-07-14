#!/usr/bin/env python

import bpy
import logging
import sys
import argparse
import os

import pathlib

logger = logging.getLogger(__name__)

texture_map = {
    'Top Copper': 'F_Cu',
    'Bottom Copper': 'B_Cu',
    'Top Silkscreen': 'F_SilkS',
    'Bottom Silkscreen': 'B_SilkS',
    'Top Soldermask': 'F_Mask',
    'Bottom Soldermask': 'B_Mask',
}


def get_by_label(nodes, label):
    for node in nodes:
        if node.label == label:
            return node
    raise Exception("Node %s not found" % label)


def instantiate_template(output_file, template_file, wrl, texture_directory):
    logger.debug("Open template file: %s" % template_file)
    bpy.ops.wm.open_mainfile(filepath=template_file)

    logger.debug("Create new pcb_collection")
    pcb_collection = bpy.data.collections.new("PCB")
    bpy.context.scene.collection.children.link(pcb_collection)

    logger.debug("Import WRL file: %s" % wrl)
    if not os.path.exists(wrl):
        raise Exception("WRL file doesn't exist")
    bpy.ops.import_scene.x3d(filepath=wrl)
    for obj in bpy.context.selected_objects:
        for coll in obj.users_collection:
            # Unlink the object
            coll.objects.unlink(obj)
        pcb_collection.objects.link(obj)
        obj.select_set(False)

    pcb_collection = bpy.data.collections.get("PCB")

    logger.debug("Merge PCB object (PCB edges and holes, with surfaces one)")
    lasts = pcb_collection.all_objects[-2:]
    for last in lasts:
        last.select_set(True)
        bpy.context.view_layer.objects.active = last
    bpy.ops.object.join()
    for obj in bpy.context.selected_objects:
        obj.select_set(False)

    logger.debug("Link PCB material with textures")
    pcb_mat = bpy.data.materials.get("PCB")
    for node_label, file_pattern in texture_map.items():
        svg_files = list(pathlib.Path(texture_directory).glob('*' + file_pattern + '.png'))
        if len(svg_files) != 1:
            logger.warning("File for pattern \"%s\" not found" % file_pattern)
            continue
        get_by_label(pcb_mat.node_tree.nodes, node_label).image = bpy.data.images.load(filepath=str(svg_files[0]))

    logger.debug("Create UV Maps")
    pcb_mesh = pcb_collection.all_objects[-1]
    top_map = pcb_mesh.data.uv_layers.new(name='Top UV Map')
    bottom_map = pcb_mesh.data.uv_layers.new(name='Bottom UV Map')
    get_by_label(pcb_mat.node_tree.nodes, "UV Map Top").uv_map = top_map.name
    get_by_label(pcb_mat.node_tree.nodes, "UV Map Bottom").uv_map = bottom_map.name

    logger.debug("Add PCB material to PCB Mesh")
    pcb_mesh.data.materials[0] = pcb_mat

    logger.debug("Save final file: %s" % output_file)
    bpy.ops.wm.save_as_mainfile(filepath=output_file)


def _main(argv=sys.argv):
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(threadName)s - %(name)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(prog=argv[0], description='KicadBlender',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-v', '--verbose', dest='verbose_count', action='count', default=0,
                        help="increases log verbosity for each occurrence.")
    parser.add_argument('template',
                        help="template file")
    parser.add_argument('wrl',
                        help="wrl input file")
    parser.add_argument('textures',
                        help="textures input directory")
    parser.add_argument('output',
                        help="input blender file")

    # Parse
    args, unknown_args = parser.parse_known_args(argv[1:])

    # Set logging level
    logger.setLevel(max(3 - args.verbose_count, 0) * 10)

    instantiate_template(args.output, args.template, args.wrl, args.textures)


def main():
    try:
        sys.exit(_main(sys.argv[sys.argv.index("--"):]))
    except Exception as e:
        logger.exception(e)
        sys.exit(-1)
    finally:
        logging.shutdown()


if __name__ == "__main__":
    main()
