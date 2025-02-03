import json
from typing import Any
from xml.etree.ElementTree import ElementTree
import easygui
import ezdxf
from ezdxf.math import Vec2
from ezdxf.render import ARROWS
import csv
import re

namespace = '{http://trimble.com/schema/fxl}'
doc = ezdxf.new(dxfversion="R2010")
msp = doc.modelspace()

# this dict is based on the colours selected the feature definition library.
# Maps Trimble Library colours to Autocad colour
MAP_RGB_TO_CAD = {
    'FF804000': '36',	#	(128, 64, 0)
    'FFD3D3D3': '254',	    #	(211, 211, 211)
    'FFFFFF00': '2', 	#	(255, 255, 0)
    'FF008000': '96', 	#	(0, 128, 0)
    'FF0080FF': '150', 	#	(0, 128, 255)
    'FFFF7F00': '30', 	#	(255, 127, 0)
    'FFFF003F': '240', 	#	(255, 0, 63)
    'FF000040': '178', 	#	(0, 0, 64)
    'FFFFA500': '40', 	#	(255, 165, 0)
    'FFADADAD': '253', 	#	(173, 173, 173)
    'FFFF00FF': '6', 	#	(255, 0, 255)
    'FF3F00FF': '180', 	#	(63, 0, 255)
    'FF800080': '216', 	#	(128, 0, 128)
    'FF0000A0': '174',     #   (0, 0, 160)
    'FF0000FF': '5',     #   (0, 0, 255)
    'FF003FFF': '160', 	#	(0, 63, 255)
    'FF009926': '104', 	#	(0, 153, 38)
    'FF00FF00': '3',  	#	(0, 255, 0)
    'FFFF0000': '1', 	#	(255, 0, 0)
    'FFFFFFFF': '7' 	    #	(255, 255, 255)
}

TEXT_HEIGHT = 0.01
LINE_FACTOR = 1.4
LEADER_LINE = Vec2(1, .2)
BLOCK_TEXT_HEIGHT = 0.08
OFFSET_RIGHT = 0.005
OFFSET_UP = 0.004


M_LEADER_LANDING_GAP = 0.4
M_LEADER_DOGLEG_LENGTH = 0.5
M_LEADER_ARROW_SIZE = 0.2
M_LEADER_SCALE = 1
M_LEADER_TEXT_HEIGHT = 0.3
M_LEADER_GAP = 0.1

# What text do you want beside a recorded point depending on its point_code
POINTCODE_TO_TEXT = {
    'wmf_av': 'A.V.',
    'wmf_bb': 'BB',
    'wmf_bm': 'B.M.',
    'wmf_conn': 'Connection\nEasting: ###\nNorthing: ###',
    'wmf_dc': 'D.C',
    'wmf_fh': 'F.H.',
    'wmf_fm': 'FM',
    'wmf_sv': 'S.V',
    'wmf_marker': 'MarkerPlate',
    'wmf_sc': 'S.C.',
    'wmf_scv': 'Sc.V',
    'wmf_scmh': 'Sc.MH',
    'wmf_rd': 'RD',
    'wmf_sd': 'S.D.',
    'wmf_tp': 'Tee',
}


# this adds a multileader with a single arrow/text.
def create_multileader_with_text(msp_ref, location, point_code):
    lead_text = ""
    if point_code in POINTCODE_TO_TEXT:
        lead_text = POINTCODE_TO_TEXT[point_code]

    for angle in [15]:
        ml_builder = msp_ref.add_multileader_mtext("EZDXF")
        ml_builder.set_connection_properties(landing_gap=M_LEADER_LANDING_GAP,
                                             dogleg_length=M_LEADER_DOGLEG_LENGTH)
        ml_builder.quick_leader( f"{lead_text}",
            target=location,
            segment1=Vec2.from_deg_angle(angle, 1),
        )
        ml_builder.set_arrow_properties(name=ARROWS.closed_blank, size=M_LEADER_ARROW_SIZE)


def create_block(block_name, block_text):
    """
    Creates and inserts a new DXF block into the document object.

    :param block_name: The name of the block.
    :param block_text: Text to be inserted into the block
    """
    global doc
    try:
        block = doc.blocks.new(name=block_name)
    except ezdxf.DXFError:
        print(f"Block '{block_name}' already exists.")
        block = doc.blocks.get(block_name)

    # Add a circle
    center = (0, 0)
    radius = .1
    block.add_circle(center, radius)

    # Add the cross (horizontal and vertical lines)
    block.add_line(start=(-radius, 0), end=(radius, 0))  # Horizontal line
    block.add_line(start=(0, -radius), end=(0, radius))  # Vertical line

    # Add text at (9, 9) with bottom-left justification
    blocktext = block.add_text(
        text=block_text,
        dxfattribs={'height': BLOCK_TEXT_HEIGHT},  # Text height
    )
    blocktext.dxf.insert = (.09, .09)
    blocktext.dxf.halign = 0  # Horizontal: Left
    blocktext.dxf.valign = 1  # Vertical: Bottom
    # Add correctly hatched quarters
    hatch = block.add_hatch(color=7)  # Default white color (adjust as needed)

    # Quarter 1 (0° to 90°)
    path1 = hatch.paths.add_edge_path()
    path1.add_line(start=center, end=(radius, 0))  # Line from center to 0°
    path1.add_arc(center=center, radius=radius, start_angle=0, end_angle=90)  # Arc from 0° to 90°
    path1.add_line(start=(0, radius), end=center)  # Line back to center

    # Quarter 3 (180° to 270°)
    path2 = hatch.paths.add_edge_path()
    path2.add_line(start=center, end=(-radius, 0))  # Line from center to 180°
    path2.add_arc(center=center, radius=radius, start_angle=180, end_angle=270)  # Arc from 180° to 270°
    path2.add_line(start=(0, -radius), end=center)  # Line back to center


def insert_block(msp_ref, block_name, locations):
    for location in locations:
        msp_ref.add_blockref(block_name, location)


def get_layers(file_path):
    """
    Parses an XML file and extracts layer definitions into a dictionary.

    This function reads a provided XML file, searches for specific
    LayerDefinitions within it, and converts each layer's data
    into a structured dictionary format. Each layer is identified by
    its name, and associated properties are stored as key-value pairs.

    :param file_path: Path to the XML file to be parsed.
    :type file_path: str
    :return: A dictionary where the keys are layer names, and the values
        are dictionaries containing each layer's properties and their values.
    :rtype: Dict[str, Dict[str, Optional[str]]]
    """
    layers_ref = {}
    tree = ElementTree()
    tree.parse(source=file_path)
    root = tree.getroot()
    for child in root:
        if f'{namespace}LayerDefinitions' in child.tag:
            for layer in child:
                if f'{namespace}LayerDefinition' in layer.tag:
                    layer_values = {}
                    name_value = "x"
                    for sub_elem in layer:
                        layer_property = sub_elem.tag.split("}")[1]
                        # name_value = ""
                        if layer_property == 'Name':
                            name_value = sub_elem.text
                            layers_ref.setdefault(name_value, {})
                        else:
                            layer_values.setdefault(layer_property, sub_elem.text)
                    layers_ref[name_value].update(layer_values)
    return layers_ref


def process_code(input_string: str, line_codes_ref, point_codes_ref) -> (bool, str, int):
    pattern = r"([a-zA-Z]+[a-z_A-Z0-9]*[a-zA-Z])(\d+)?$"
    found = False
    point_type = "Unknown" # default if your pointcode is not defined in Library
    match = re.match(pattern, input_string)
    code = False # the point code extracted from the (from kb3 => get kb)
    differentiator = False # the number of this pointcode, for separating lines (from kb3 => get 3)
    if match:
        code = match.group(1)  # Letters or alphanumeric code
        differentiator = match.group(2)  # Number at the end
        if code in line_codes_ref:
            found = True
            point_type = "Line"
        if code in point_codes_ref:
            found = True
            point_type = "Point"
        if differentiator:
            differentiator = int(differentiator)
        else:
            if code in line_codes_ref:
                differentiator = -1
    else:
        print("The string does not match the expected format.")
    if code and found:
        print(f"code:{code}, differentiator:{differentiator}, point_type:{point_type}")
        return True, code, differentiator, point_type
    else:
        return False,"sprl", 1, point_type


def get_codes(file_path):
    """
    Extracts point and line feature definitions and their attributes from a Trimble Library FXL file.

    :param file_path: Path to the XML file that contains the feature definitions.
    :type file_path: str
    :return: A tuple containing three dictionaries: the first for point feature
             definitions and the second for line feature definitions, the third maps pointcodes to their target layers.
    :rtype: Tuple[Dict[str, Dict[str, str]], Dict[str, Dict[str, str]], Dict[str, str]]
    """
    point_code_dict_ref = {}
    line_code_dict_ref = {}
    code_layer_map_ref = {}
    tree = ElementTree()
    tree.parse(file_path)
    root = tree.getroot()
    for child in root:
        if f'{namespace}FeatureDefinitions' in child.tag:
            for gchild in child:
                if f'{namespace}PointFeatureDefinition' in gchild.tag:
                    point_code = gchild.attrib.get('Code')
                    point_code_dict_ref.setdefault(point_code, gchild.attrib)
                    code_layer_map_ref.setdefault(point_code, gchild.attrib.get('Layer'))
                if f'{namespace}LineFeatureDefinition' in gchild.tag:
                    line_code = gchild.attrib.get('Code')
                    line_code_dict_ref.setdefault(line_code, gchild.attrib)
                    code_layer_map_ref.setdefault(line_code, gchild.attrib.get('Layer'))
    return point_code_dict_ref, line_code_dict_ref, code_layer_map_ref


def get_survey(file_path, point_codes_ref, line_codes_ref):
    global needed_pcodes
    fixed_columns = ["PointNumber", "Easting", "Northing", "Height", "Point_code"]

    # Open the CSV file
    try:
        with open(file_path, mode='r', encoding='utf-8', newline='') as csv_file:
            reader = csv.reader(csv_file)

            all_pts_map = {"Line":{}, "Point":{}, "Unknown":{}}
            for row in reader:
                row_dict: dict[str, Any] = {fixed_columns[i]: row[i] for i in range(min(len(row), 5))} # map <=5 cols
                point_code = row_dict.get("Point_code", False)
                if point_code: # the first 2 csv rows and further temp stations have no point code, so check it's real
                    success, pcode, sequence, point_type = process_code(row_dict['Point_code'], line_codes_ref, point_codes_ref)
                    row_dict["Point_code"] = f"{pcode}" # add these to row_dict so we can find required layer later
                    row_dict["Sequence"] = f"{sequence}"
                    if pcode not in needed_pcodes:
                        needed_pcodes.append(pcode)
                    if len(row) > 5: # anything after col5 are attribute records
                        attributes = row[5:]
                        attributes_dict = {}
                        # Process key-value pairs or just (key: "")
                        for j in range(0, len(attributes), 2):
                            key = attributes[j]
                            if j + 1 < len(attributes):  # check if a value exists for the key or make one
                                value = attributes[j + 1]
                            else:
                                value = ""
                            attributes_dict[key] = value

                        row_dict["attrib"] = attributes_dict
                    # example structure all_pts_map['Line']['bb']['1']=[{point0 details},{point2 details}]  list of dicts
                    all_pts_map[point_type].setdefault(row_dict['Point_code'], {}).setdefault(sequence, []).append(row_dict)
            return all_pts_map
    except FileNotFoundError:
        print(f"Error: Survey File '{file_path}' not found.")
        return None
    except PermissionError:
        print(f"Error: Permission denied to access survey file '{file_path}'.")
        return None


def add_attrib_text(point, point_code , layer_name="Points"):
    global msp # get the global modelspace reference
    x = float(point.get("Easting", 0))
    y = float(point.get("Northing", 0))
    z = float(point.get("Height", 0))

    if "attrib" in point: # if there are attributes, add a text item for each one, spaced rising up to right of point
        count = 0
        for attrib_key, attrib_value in point["attrib"].items():
            msp.add_mtext(
                text=f"\\H0.5x;{attrib_key.split(':')[1]}: \\H1x;{attrib_value}",
                dxfattribs={
                    "insert": (x + OFFSET_RIGHT, y + OFFSET_UP + TEXT_HEIGHT * count * LINE_FACTOR),
                    "char_height": TEXT_HEIGHT,
                    "attachment_point": 7,
                }
            )
            count += 1
        print(point["attrib"])
    msp.add_point((x, y, z), dxfattribs={"layer": f"points_{layer_name}"})
    sequence_text = point['Sequence'] if 'Sequence' in point else ''
    if sequence_text == "None" or sequence_text == "-1":
        sequence_text = ""
    msp.add_text(text=f"Code: {point_code}{sequence_text}",
                 dxfattribs={
                     "insert": (x + OFFSET_RIGHT, y - 1 * TEXT_HEIGHT * LINE_FACTOR),  # Insertion point (x, y)
                     "height": TEXT_HEIGHT,
                     "rotation": 0,
                     "layer": "Point Code"
                 })
    msp.add_text(text=f"Z = {point["Height"]}",
                 dxfattribs={
                     "insert": (x + OFFSET_RIGHT, y - 2 * TEXT_HEIGHT * LINE_FACTOR),  # Insertion point (x, y)
                     "height": TEXT_HEIGHT,
                     "rotation": 0,
                     "layer": "Point Height"
                 })
    msp.add_text(text=f"PtNum: {point["PointNumber"]}",
                 dxfattribs={
                     "insert": (x + OFFSET_RIGHT, y - 3 * TEXT_HEIGHT * LINE_FACTOR),
                     # Insertion point (x, y)
                     "height": TEXT_HEIGHT,
                     "rotation": 0,
                     "layer": "Point Number"
                 })


def create_blocks():
    global needed_pcodes
    # create_block(block_name='AV_Block', block_text='AV')
    for needed_pcode in needed_pcodes:
        if needed_pcode in POINTCODE_TO_TEXT:
            block_text = POINTCODE_TO_TEXT[needed_pcode]
            block_name = f"{block_text}_Block"
            create_block(block_name=block_name, block_text=block_text)


def create_dxf(output_dxf_path, layers, survey, point_codes_dict, line_codes_dict, needed_layers):
    doc.header["$PDMODE"] = 34  # sets point style
    doc.header["$PDSIZE"] = 0.2 # sets point size as fixed
    if "OpenSans" not in doc.styles:
        doc.styles.add("OpenSans", font="OpenSans.ttf")

    #create_block(block_name='AV_Block', block_text='AV')
    create_blocks()
    mleaderstyle = doc.mleader_styles.duplicate_entry("Standard", "EZDXF")
    mleaderstyle.set_mtext_style("OpenSans")
    mleaderstyle.dxf.scale= M_LEADER_SCALE
    mleaderstyle.dxf.char_height = M_LEADER_TEXT_HEIGHT
    mleaderstyle.dxf.landing_gap_size = M_LEADER_GAP  # Adjust gap between the leader line and the text

    # msp = doc.modelspace()
    doc.layers.add("Point Code", color=1)  # Create a new layer for points
    doc.layers.add("Point Height", color=5)  # Create a new layer for points
    doc.layers.add("Point Number", color=3)  # Create a new layer for points
    layer_color = 7
    for layer in needed_layers:
        if layer in layers.keys():
            if layer not in ["0", "Points"]:
                layer_color = MAP_RGB_TO_CAD.get(layers[layer]['Color'], 7)
        else:
            layer_color = 7

        doc.layers.add(layer, linetype="Continuous", color=int(layer_color))

    # start processing survey lines to dxf

    for point_code in survey['Line']:
        print(point_code)
        vertex_list = []
        layer_name = line_codes_dict[point_code]['Layer']
        for sequence in survey['Line'][point_code]:
            vertex_list = []
            for point in survey['Line'][point_code][sequence]:
                add_attrib_text(point, point_code, layer_name=layer_name)
                vertex_list.append((float(point.get("Easting", 0)), float(point.get("Northing", 0)), float(point.get("Height", 0))))
            msp.add_polyline3d(vertex_list, dxfattribs={"layer": layer_name})

    for point_code in survey['Unknown']:
        print(f"Unknown point code: {point_code}")
        vertex_list = []
        layer_name = line_codes_dict[point_code]['Layer']
        for sequence in survey['Unknown'][point_code]:
            vertex_list = []
            for point in survey['Unknown'][point_code][sequence]:
                add_attrib_text(point, point_code, layer_name=layer_name)
                vertex_list.append((float(point.get("Easting", 0)), float(point.get("Northing", 0)), float(point.get("Height", 0))))
            msp.add_polyline3d(vertex_list, dxfattribs={"layer": layer_name})

    # next we want to put points for each point on the points layer
    for point_code in survey['Point']:
        locations = []
        needs_block = False
        if point_code in POINTCODE_TO_TEXT:
            needs_block = True
            block_name = f"{POINTCODE_TO_TEXT[point_code]}_Block"
        print(point_code)
        layer_name = point_codes_dict[point_code]['Layer']
        for key, value in survey['Point'][point_code].items():
            for point in value:
                if needs_block:
                    locations.append((float(point.get("Easting", 0)), float(point.get("Northing", 0)), float(point.get("Height", 0))))
                    lead_point = Vec2(float(point.get("Easting", 0)), float(point.get("Northing", 0)))
                    create_multileader_with_text(msp, lead_point, point_code=point_code)
                add_attrib_text(point, point_code, layer_name=layer_name)
                if needs_block:
                    insert_block(msp, block_name=block_name, locations=locations)
    doc.saveas(output_dxf_path)


def get_layer_from_pcode(pcode_list, code_layer_map):
    needed_layers = set()
    needed_layers.add("Points")
    for pcode in pcode_list:
        if pcode in code_layer_map.keys():
            layer = code_layer_map[pcode]
            needed_layers.add(layer)
            needed_layers.add("attrib_"+layer)
            needed_layers.add("points_" + layer)
        else:
            needed_layers.add("Spare")
            needed_layers.add("attrib_Spare")
            needed_layers.add("points_Spare")
    return needed_layers


def set_dxf_view(survey_ref):
    global doc
    count = 0
    min_x = min_y = max_x = max_y = 0
    for point_code in survey_ref['Line']:
        for sequence in survey_ref['Line'][point_code]:
            for point in survey_ref['Line'][point_code][sequence]:
                x = float(point.get("Easting", 0))
                y = float(point.get("Northing", 0))
                if count == 0:
                    min_x = x
                    min_y = y
                    max_x = x
                    max_y = y
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)
                count += 1
    doc.header['$LIMMIN'] = (min_x - 20, min_y- 20)  # Lower-left corner of limits
    doc.header['$LIMMAX'] = (max_x + 20, max_y + 20)

    vp_height = max_y - min_y + 20
    vp_width = max_x - min_x + 20
    cen_x = (max_x + min_x) / 2
    cen_y = (max_y + min_y) / 2
    zoom_all = 2  # desired aspect ratio of width/height for view box, need incase survey is wide but not high
    survey_aspect = vp_width / vp_height
    vp_heightrect = vp_height
    if survey_aspect > zoom_all:
        vp_heightrect = vp_width / zoom_all
    doc.set_modelspace_vport(height=vp_heightrect , center=(cen_x, cen_y))

def get_config(config_file):
    with open(config_file, 'r') as f:
        config = json.load(f)
    return config

def save_config(config_file_ref, config_ref):
    with open(config_file_ref, 'w') as fh:
        json_string = json.dumps(config_ref, indent=4)  # Serialize dictionary to JSON string
        with open(config_file_ref, 'w') as fh:  # Open file for writing
            fh.write(json_string)  # Write string to the file


if __name__ == "__main__":
    if __debug__:
        library_file_path = "Global_v3.fxl"
        survey_file_path = "tr3.csv"
        output_dxf_path = "Output6.dxf"
    else:
        config_file = "./config.json"
        config = get_config(config_file)

        library_file_path = easygui.fileopenbox(
            title="Select Library File",
            filetypes=["*.fxl"],
            default=config.get("library_file",  "Global_v2.fxl")
        )
        config["library_file"] = library_file_path

        survey_file_path = easygui.fileopenbox(
            title="Select Survey File",
            filetypes=["*.csv"],
            default=config.get("survey_file", "tr3.csv")
        )
        config["survey_file"] = survey_file_path

        output_dxf_path = easygui.filesavebox(
            title="Select Output DXF File",
            filetypes=["*.dxf"],
            default=config.get("output_dxf", "output.dxf")
        )
        config["output_dxf"] = output_dxf_path

        save_config(config_file, config)

    needed_pcodes = []
    # Get info from FXL file
    layers = get_layers(library_file_path)
    point_code_dict, line_code_dict, code_layer_map = get_codes(library_file_path)
    point_codes = list(point_code_dict.keys())
    line_codes = list(line_code_dict.keys())

    # Get survey using Line and Point Codes
    survey = get_survey(survey_file_path, point_codes, line_codes)

    # Set up the DXF view based on survey coordinates
    set_dxf_view(survey)

    # find which layers need to be inserted in the DXF and the attributes they require

    needed_layers = get_layer_from_pcode(needed_pcodes, code_layer_map)

    # add survey items to DXF
    create_dxf(output_dxf_path, layers, survey, point_code_dict, line_code_dict, needed_layers)
