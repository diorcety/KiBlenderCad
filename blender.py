#!/usr/bin/env python

import os
import sys
# cx_freeze issue like this one https://github.com/pyinstaller/pyinstaller/issues/3795
if sys.platform == "win32":
    import ctypes
    ctypes.windll.kernel32.SetDllDirectoryA(None)


import bpy
import bmesh

import logging
import argparse
import pathlib

from mathutils.bvhtree import BVHTree
from collections import defaultdict
from mathutils import Vector, Quaternion, Matrix
from functools import reduce
from itertools import product

logger = logging.getLogger(__name__)


###
### FROM https://gist.github.com/SURYHPEZ/9502819
###

def merge_boxes(objects):
    return reduce(lambda x, y: x + y, [Box(obj) for obj in objects if obj.type == 'MESH'])


class Box:
    def __init__(self, bl_object=None, max_min=None):
        if bl_object and bl_object.type == 'MESH':
            self.__bound_box = self.__get_bound_box_from_object(bl_object)
        elif max_min:
            self.__bound_box = self.__get_bound_box_from_max_min(max_min)
        else:
            raise TypeError()

    def __add__(self, bound_box):
        return self.merge(bound_box)

    def __getitem__(self, index):
        return self.__bound_box[index]

    @property
    def max(self):
        return Vector(max((v.x, v.y, v.z) for v in self.__bound_box))

    @property
    def min(self):
        return Vector(min((v.x, v.y, v.z) for v in self.__bound_box))

    @property
    def center(self):
        return sum((v for v in self.__bound_box), Vector()) / 8

    def merge(self, box):
        if not box:
            return self

        if not isinstance(box, Box):
            raise TypeError('Require a Box object')

        max_new = Vector(map(max, zip(self.max, box.max)))
        min_new = Vector(map(min, zip(self.min, box.min)))

        return Box(max_min=(max_new, min_new))

    def __get_bound_box_from_object(self, bl_object):
        return [bl_object.matrix_world @ Vector(v) for v in bl_object.bound_box]

    def __get_bound_box_from_max_min(self, max_min):
        max_point, min_point = max_min

        return [Vector(v) for v in product((max_point.x, min_point.x),
                                           (max_point.y, min_point.y),
                                           (max_point.z, min_point.z))]


###
###
###

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


def bvh_from_bmesh(obj):
    bm = bmesh.new()
    bm.from_mesh(obj.data)
    bm.transform(obj.matrix_world)
    result = BVHTree.FromBMesh(bm)
    return result, bm


def get_closest_distance(bmesh1, bmesh2, max_distance=0.01):
    min_d = float("inf")
    _, bm1 = bmesh1
    bvh2, _ = bmesh2
    for v in bm1.verts:
        r = bvh2.find_nearest_range(v.co, max_distance)
        if len(r) > 0:
            min_d = min(min([a[3] for a in r]), min_d)
    return min_d


def is_overlapped(bmesh1, bmesh2):
    bvh1, _ = bmesh1
    bvh2, _ = bmesh2
    return len(bvh1.overlap(bvh2, )) > 0


def join_tree_objects_with_tree(tree, tree_objects):
    name = tree.name
    tree_objects = list(tree_objects)
    tree_objects.insert(0, tree)
    ctx = bpy.context.copy()
    ctx["active_object"] = tree_objects[0]
    ctx["selected_editable_objects"] = tree_objects
    bpy.ops.object.join(ctx)
    return bpy.data.objects[name]


def factorize_mats(objects):
    dct = defaultdict(set)

    # Get All object mats
    for obj in objects:
        for mat_s in obj.material_slots:
            mat = mat_s.material
            if mat.use_nodes:
                continue
            identifier = (tuple(mat.diffuse_color), mat.specular_color.copy().freeze())
            dct[identifier].add(mat)

    for mat_set in dct.values():
        mat_list = list(mat_set)
        for mat in mat_list[1:]:
            mat.user_remap(mat_list[0])
            bpy.data.materials.remove(mat)


def regroup_meshs(objects, max_distance=0.05):
    initial_len = len(objects)

    objects = set(objects)
    world_v = {obj: bvh_from_bmesh(obj) for obj in objects}

    def remove_from_world_vertices(a):
        bvh, bm = world_v.pop(a)
        bm.free()

    while len(objects) > 0:
        logger.debug("Progress %d" % ((initial_len - len(objects)) / initial_len * 100))
        to_merge = set()

        obj_a = objects.pop()
        for obj_b in objects:
            if is_overlapped(world_v[obj_a], world_v[obj_b]):
                to_merge.add(obj_b)
            else:
                d = get_closest_distance(world_v[obj_a], world_v[obj_b], max_distance)
                if d <= max_distance:
                    to_merge.add(obj_b)

        remove_from_world_vertices(obj_a)

        for obj in to_merge:
            # Find object with same data (have to be really removed)
            for o in bpy.data.objects:
                if o.data == obj.data and o != obj:
                    if o in objects:
                        objects.remove(o)
                        remove_from_world_vertices(o)
                    bpy.data.objects.remove(o, do_unlink=True)
            objects.remove(obj)
            remove_from_world_vertices(obj)

        if len(to_merge) > 0:
            ret = join_tree_objects_with_tree(obj_a, to_merge)

            # Update data
            objects.add(ret)
            world_v[ret] = bvh_from_bmesh(ret)


def get_mass_center(obj):
    local_bbox_center = 0.125 * sum((Vector(b) for b in obj.bound_box), Vector())
    global_bbox_center = obj.matrix_world @ local_bbox_center
    return global_bbox_center


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


def get_pcb(objects, dimensions, approximation=0.03):
    result = []
    for o in objects:
        if sum([1 if len(list(filter(lambda x: abs(x - y) / y <= approximation, o.dimensions))) >= 1 else 0 for y in
                dimensions]) == len(dimensions):
            result.append(o)
    return result


def fancy_positioning(camera, focus, all_objects):
    origin = Vector((0, 0, 0))
    objects_box = merge_boxes(all_objects)
    dimensions = objects_box.max - objects_box.min
    direction = (camera.location - focus.location).normalized()
    axis_align = Vector(([1.0 if dimensions[i] == min(dimensions) else 0.0 for i in range(3)]))
    angle = axis_align.angle(direction)
    axis = axis_align.cross(direction)

    r = Quaternion(axis, angle).to_matrix().to_4x4()
    m1 = Matrix.Translation(origin - objects_box.center)
    m2 = Matrix.Translation(focus.location - origin)

    # Move to origin, rotate and move to focus point
    for o in all_objects:
        o.matrix_world = (m2 @ r @ m1) @ o.matrix_world


def instantiate_template(output_file, template_file, pcb_dimensions, wrl, texture_directory):
    logger.debug("Open template file: %s" % template_file)
    bpy.ops.wm.open_mainfile(filepath=template_file)

    bpy.ops.object.select_all(action='DESELECT')

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

    logger.debug("Clean up duplicated materials")
    factorize_mats(pcb_collection.all_objects)

    logger.debug("Merge PCB objects (PCB edges and holes, with surface ones)")
    pcb_objects = None
    if pcb_dimensions is not None:
        pcb_objects = get_pcb(pcb_collection.all_objects, pcb_dimensions)
        if len(pcb_objects) == 0:
            dimensions_str = (', '.join(map(str, pcb_dimensions)))
            logger.warning("Can't find objects with same dimensions as specified: %s" % dimensions_str)
    if pcb_objects is None:
        logger.warning("Use two last objects of imported VRML as PCB objects (default VRML Kicad exporter behaviour)")
        pcb_objects = pcb_collection.all_objects[-2:]
    if len(pcb_objects) > 1:
        pcb_object = join_tree_objects_with_tree(pcb_objects[0], pcb_objects[1:])
    else:
        pcb_object = pcb_objects[0]
    pcb_object.name = "Board"

    logger.debug("Merge mesh from same components")
    component_mesh = set(pcb_collection.all_objects)
    component_mesh.remove(pcb_object)
    regroup_meshs(component_mesh)

    logger.debug("Clean up duplicated vertices again")
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

    logger.debug("Positioning the board")
    fancy_positioning(bpy.data.objects['Camera'], bpy.data.objects['Camera_Focus'], pcb_collection.all_objects)

    logger.debug("Save final file: %s" % output_file)
    bpy.ops.wm.save_as_mainfile(filepath=output_file)


def dim_type(strings):
    return tuple(map(float, strings.split(":")))


def _main(argv=sys.argv):
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(threadName)s - %(name)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(prog=argv[0], description='KicadBlender',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-v', '--verbose', dest='verbose_count', action='count', default=0,
                        help="increases log verbosity for each occurrence.")
    parser.add_argument('-d', '--dimensions', type=dim_type,
                        help="pcb dimensions: width:height:tickness")
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

    instantiate_template(args.output, args.template, args.dimensions, args.wrl, args.textures)


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
