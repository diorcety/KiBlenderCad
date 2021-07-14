# KiBlenderCad
Based on the work of @PCB-Arts: https://github.com/PCB-Arts/stylized-blender-setup
Automatically create a blender file of a kicad_pcb based on a template:

* Export layers in SVG files
* (TODO Kicad version 6) Export VRML file
* Convert SVG files to PNG textures thanks to inkscape
* Integrate PNG layers and maps textures in Blender files

```
generator.py -vvvvv "D:\MyPcb\MyPcb.kicad_pcb" "D:\MyPcb\Blender"
