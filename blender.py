#!/usr/bin/env python

import bpy
import bmesh

import logging
import sys
import argparse
import os

import pathlib

from mathutils.bvhtree import BVHTree

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


def BVHFromBMesh(obj):
    bm = bmesh.new()
    bm.from_mesh(obj.data)
    bm.transform(obj.matrix_world)
    result = BVHTree.FromBMesh(bm)
    return (result, bm)


def GetClosestDistance(bmesh1, bmesh2, max_distance=0.01):
    min_d = float("inf")
    bvh1, bm1 = bmesh1
    bvh2, bm2 = bmesh2
    for v in bm1.verts:
        r = bvh2.find_nearest_range(v.co, max_distance)
        if len(r) > 0:
            min_d = min(min([a[3] for a in r]), min_d)
    return min_d


def JoinTreeObjectsWithTree(tree, treeObjects):
    name = tree.name
    treeObjects = list(treeObjects)
    treeObjects.insert(0, tree)
    ctx = bpy.context.copy()
    ctx["active_object"] = treeObjects[0]
    ctx["selected_editable_objects"] = treeObjects
    bpy.ops.object.join(ctx)
    return bpy.data.objects[name]


def RegroupMeshs(objects, max_distance=0.05):
    initial_len = len(objects)

    objects = set(objects)
    world_v = {obj: BVHFromBMesh(obj) for obj in objects}

    def removeFromWorldVertices(a):
        bvh, bm = world_v.pop(a)
        bm.free()

    while len(objects) > 0:
        logger.debug("Progress %d" % ((initial_len - len(objects)) / initial_len * 100))
        to_merge = set()

        obj_a = objects.pop()
        for obj_b in objects:
            d = GetClosestDistance(world_v[obj_a], world_v[obj_b], max_distance)
            if d <= max_distance:
                to_merge.add(obj_b)

        removeFromWorldVertices(obj_a)

        for obj in to_merge:
            # Find object with same data (have to be really removed)
            for o in bpy.data.objects:
                if o.data == obj.data and o != obj:
                    if o in objects:
                        objects.remove(o)
                        removeFromWorldVertices(o)
                    bpy.data.objects.remove(o, do_unlink=True)
            objects.remove(obj)
            removeFromWorldVertices(obj)

        if len(to_merge) > 0:
            ret = JoinTreeObjectsWithTree(obj_a, to_merge)

            # Update data
            objects.add(ret)
            world_v[ret] = BVHFromBMesh(ret)


def cleanup(objects, distance=0.000001):
    meshes = set(o.data for o in objects if o.type == 'MESH')

    bm = bmesh.new()

    for m in meshes:
        bm.from_mesh(m)
        bmesh.ops.remove_doubles(bm, verts=bm.verts, dist=distance)
        bm.to_mesh(m)
        m.update()
        bm.clear()

    bm.free()


def instantiate_template(output_file, template_file, wrl, texture_directory):
    logger.debug("Open template file: %s" % template_file)
    bpy.ops.wm.open_mainfile(filepath=template_file)

    logger.debug("Create PCB collection")
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

    logger.debug("Clean up duplicated vertices")
    cleanup(pcb_collection.all_objects)

    logger.debug("Merge PCB objects (PCB edges and holes, with surface ones)")
    lasts = pcb_collection.all_objects[-2:]
    pcb_object = JoinTreeObjectsWithTree(lasts[0], lasts[1:])

    logger.debug("Merge mesh from same components")
    component_mesh = set(pcb_collection.all_objects)
    component_mesh.remove(pcb_object)
    RegroupMeshs(component_mesh)

    logger.debug("Clean up duplicated vertices")
    cleanup(pcb_collection.all_objects)

    logger.debug("Link PCB material with textures")
    pcb_mat = bpy.data.materials.get("PCB")
    for node_label, file_pattern in texture_map.items():
        svg_files = list(pathlib.Path(texture_directory).glob('*' + file_pattern + '.png'))
        if len(svg_files) != 1:
            logger.warning("File for pattern \"%s\" not found" % file_pattern)
            continue
        get_by_label(pcb_mat.node_tree.nodes, node_label).image = bpy.data.images.load(filepath=str(svg_files[0]))

    logger.debug("Create UV Maps")
    top_map = pcb_object.data.uv_layers.new(name='Top UV Map')
    bottom_map = pcb_object.data.uv_layers.new(name='Bottom UV Map')
    get_by_label(pcb_mat.node_tree.nodes, "UV Map Top").uv_map = top_map.name
    get_by_label(pcb_mat.node_tree.nodes, "UV Map Bottom").uv_map = bottom_map.name

    logger.debug("Add PCB material to PCB Mesh")
    pcb_object.data.materials[0] = pcb_mat

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
