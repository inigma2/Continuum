import zipfile
import re
import io
from collections import defaultdict
import math
import os
import sys
import shutil

# Conditional import for Windows-specific registry access
if sys.platform == "win32":
    import winreg

# --- CONFIGURATION ---
SUPPORTED_STELLARIS_VERSION = "4.0"

def clear_screen():
    """Clears the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def find_stellaris_user_dir():
    """Finds the Stellaris user documents directory by querying the OS directly."""
    if sys.platform == "win32":
        # More robust method using "User Shell Folders" which handles OneDrive redirection
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
                doc_path = winreg.QueryValueEx(key, "Personal")[0]
            # Paths from this key can have environment variables like %USERPROFILE%
            doc_path = os.path.expandvars(doc_path)
            stellaris_dir = os.path.join(doc_path, 'Paradox Interactive', 'Stellaris')
            if os.path.isdir(stellaris_dir):
                return stellaris_dir
        except Exception:
            pass  # Silently fail and try the next method

        # Original method as a fallback
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                doc_path = winreg.QueryValueEx(key, "Personal")[0]
            stellaris_dir = os.path.join(doc_path, 'Paradox Interactive', 'Stellaris')
            if os.path.isdir(stellaris_dir):
                return stellaris_dir
        except Exception:
            print("Warning: Could not query Windows Registry for Documents path. Using standard fallback.")

        # Final fallback for non-standard setups
        doc_path = os.path.join(os.path.expanduser('~'), 'Documents')
        return os.path.join(doc_path, 'Paradox Interactive', 'Stellaris')

    elif sys.platform == "darwin":
        return os.path.join(os.path.expanduser('~'), 'Documents', 'Paradox Interactive', 'Stellaris')
    elif sys.platform == "linux" or sys.platform == "linux2":
        return os.path.join(os.path.expanduser('~'), '.local', 'share', 'Paradox Interactive', 'Stellaris')
    return None

def find_stellaris_install_dir():
    """Finds the Stellaris game install directory by checking Steam's library files."""
    steam_path = ""
    if sys.platform == "win32":
        try:
            hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Valve\\Steam")
            steam_path = winreg.QueryValueEx(hkey, "InstallPath")[0]
        except FileNotFoundError:
            try:
                hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Valve\\Steam")
                steam_path = winreg.QueryValueEx(hkey, "InstallPath")[0]
            except FileNotFoundError: return None
    elif sys.platform == "darwin":
        steam_path = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', 'Steam')
    elif sys.platform == "linux" or sys.platform == "linux2":
        steam_path = os.path.join(os.path.expanduser('~'), '.steam', 'steam')

    if not steam_path or not os.path.isdir(steam_path): return None
    library_folders_file = os.path.join(steam_path, 'steamapps', 'libraryfolders.vdf')
    if not os.path.exists(library_folders_file): return None
    library_paths = [os.path.join(steam_path)]
    try:
        with open(library_folders_file, 'r') as f:
            for line in f:
                match = re.search(r'"path"\s+"([^"]+)"', line)
                if match: library_paths.append(match.group(1).replace('\\\\', '\\'))
    except Exception as e:
        print(f"Warning: Could not parse Steam library folders file: {e}")
    for path in library_paths:
        stellaris_path = os.path.join(path, 'steamapps', 'common', 'Stellaris')
        if os.path.isdir(stellaris_path): return stellaris_path
    return None

def get_save_meta_data(save_file_path):
    """Reads the version and date from the meta file inside a .sav archive."""
    version = "Unknown"
    date = "Unknown Date"
    try:
        with zipfile.ZipFile(save_file_path, 'r') as save_zip:
            if 'meta' in save_zip.namelist():
                with save_zip.open('meta') as meta_file:
                    meta_content = io.TextIOWrapper(meta_file, encoding='utf-8').read()
                    version_match = re.search(r'version="([^"]+)"', meta_content)
                    if version_match: version = version_match.group(1)
                    date_match = re.search(r'date="([^"]+)"', meta_content)
                    if date_match: date = date_match.group(1)
    except Exception as e:
        print(f"Warning: Could not read metadata for {os.path.basename(save_file_path)}. {e}")
    return version, date

def get_stellaris_language(user_dir):
    settings_path = os.path.join(user_dir, 'settings.txt')
    try:
        with open(settings_path, 'r', encoding='utf-8') as f:
            match = re.search(r'language="(\w+)"', f.read())
            if match:
                language_key = match.group(1).lstrip('l_')
                print(f"Detected language: {language_key}")
                return language_key
    except Exception:
        print("Warning: Could not detect language. Defaulting to English.")
    return "english"

def load_localization_data(install_dir, language):
    localization_map = {}
    base_loc_path = os.path.join(install_dir, 'localisation', language)
    print(f"Searching for all localization files in:\n{base_loc_path}\n")
    if not os.path.isdir(base_loc_path): return {}
    loc_pattern = re.compile(r'([\w_.-]+):\d*\s*"(.*?)"')
    for root, _, files in os.walk(base_loc_path):
        for filename in files:
            if filename.endswith(f'l_{language}.yml'):
                file_path = os.path.join(root, filename)
                try:
                    with open(file_path, 'r', encoding='utf-8-sig') as f:
                        content = f.read()
                        matches = loc_pattern.findall(content)
                        for key, value in matches:
                            localization_map[key] = value
                except Exception as e:
                    print(f"Warning: Error reading file {file_path}: {e}")
    print(f"Loaded {len(localization_map)} localization keys.")
    return localization_map

def resolve_name(name_block_content, loc_data, star_count_context=None, parent_body_name=None):
    if not name_block_content: return "Unknown"
    key_match = re.search(r'^\s*key="([^"]+)"', name_block_content)
    if not key_match:
        key_match_simple = re.search(r'key="([^"]+)"', name_block_content)
        if not key_match_simple: return "Unknown"
        name_key = key_match_simple.group(1)
    else:
        name_key = key_match.group(1)

    if name_key.startswith('$') and name_key.endswith('$'): name_key = name_key.strip('$')
    variables_content = _get_nested_block_content(name_block_content, r'variables\s*=\s*{')
    if (name_key.startswith("STAR_NAME_") or name_key.endswith("_NAME_FORMAT") or name_key.startswith("NEW_COLONY_NAME")) and variables_content:
        if name_key.startswith("STAR_NAME_") or name_key.startswith("NEW_COLONY_NAME"):
            name_value_block = _get_nested_block_content(variables_content, r'key="NAME"\s*value\s*=\s*{')
            if name_value_block:
                base_name = resolve_name(name_value_block, loc_data, star_count_context)
                if name_key.startswith("NEW_COLONY_NAME"): return f"{base_name} Prime"
                star_match = re.match(r'STAR_NAME_(\d)_OF_(\d)', name_key)
                if star_match and star_count_context is not None and star_count_context > 1:
                    num = int(star_match.group(1))
                    if 1 <= num <= 3: return f"{base_name} {('ABC')[num - 1]}"
                return base_name
            return "Unknown Star"
        if name_key == "PLANET_NAME_FORMAT":
            parent_value_block = _get_nested_block_content(variables_content, r'key="PARENT"\s*value\s*=\s*{')
            numeral_value_block = _get_nested_block_content(variables_content, r'key="NUMERAL"\s*value\s*=\s*{')
            if parent_value_block and numeral_value_block:
                parent_name_val = resolve_name(parent_value_block, loc_data, star_count_context)
                numeral_key_match = re.search(r'key="([^"]+)"', numeral_value_block)
                if numeral_key_match: return f"{parent_name_val} {numeral_key_match.group(1)}"
            return "Unknown Planet"
        if name_key == "SUBPLANET_NAME_FORMAT":
            parent_value_block = _get_nested_block_content(variables_content, r'key="PARENT"\s*value\s*=\s*{')
            numeral_matches = re.findall(r'key="NUMERAL"\s*value\s*=\s*{\s*key="([^"]+)"', variables_content, re.DOTALL)
            if parent_value_block and numeral_matches:
                moon_base_name = resolve_name(parent_value_block, loc_data, star_count_context)
                moon_numeral = numeral_matches[-1]
                if parent_body_name and moon_base_name != parent_body_name: return moon_base_name
                is_roman_planet_numeral = (len(moon_numeral) > 1) or (moon_numeral.upper() in ['I', 'V', 'X'])
                
                if is_roman_planet_numeral:
                    return f"{moon_base_name} {moon_numeral}"
                else:
                    return f"{moon_base_name}{moon_numeral.lower()}"
            return "Unknown Moon"
        if name_key == "ASTEROID_NAME_FORMAT":
            prefix, suffix = "",""
            prefix_val_block = _get_nested_block_content(variables_content, r'key="prefix"\s*value\s*=\s*{')
            if prefix_val_block:
                prefix_match = re.search(r'key="([^"]+)"', prefix_val_block)
                if prefix_match: prefix = prefix_match.group(1)
            suffix_val_block = _get_nested_block_content(variables_content, r'key="suffix"\s*value\s*=\s*{')
            if suffix_val_block:
                suffix_match = re.search(r'key="([^"]+)"', suffix_val_block)
                if suffix_match: suffix = suffix_match.group(1)
            return f"{prefix}{suffix}"
    if name_key in loc_data: return loc_data[name_key]
    clean_name = re.sub(r'(_system|_SYSTEM)$', '', name_key)
    clean_name = re.sub(r'^(NAME_|SPEC_)', '', clean_name)
    return clean_name.replace('_', ' ')

def _get_nested_block_content(text, start_regex):
    match = re.search(start_regex, text)
    if not match: return None
    content_start_index = match.end()
    brace_level = 1
    for i in range(content_start_index, len(text)):
        char = text[i]
        if char == '{': brace_level += 1
        elif char == '}': brace_level -= 1
        if brace_level == 0: return text[content_start_index:i]
    return None

def build_galaxy_hierarchy(stars, planets, loc_data):
    moons_by_parent = defaultdict(list)
    for planet_id, planet_data in planets.items():
        if 'moon_of' in planet_data: moons_by_parent[planet_data['moon_of']].append(planet_data)
    hierarchical_systems = []
    for star_id, star_data in stars.items():
        system = star_data; system['planets'] = []
        system['system_star_class'] = system.get('star_class', 'sc_g')
        if 'raw_name_block' in star_data: system['name'] = resolve_name(star_data['raw_name_block'], loc_data)
        for planet_id in star_data.get('planet_ids', []):
            if planet_id in planets:
                planet_data = planets[planet_id]
                if 'moon_of' not in planet_data:
                    if 'x' in planet_data and 'y' in planet_data: planet_data['orbit'] = math.sqrt(float(planet_data['x'])**2 + float(planet_data['y'])**2)
                    elif 'orbit' in planet_data: planet_data['orbit'] = abs(float(planet_data['orbit']))
                    system['planets'].append(planet_data)
        def attach_moons_recursively(body_list):
            for body in body_list:
                if body.get('id') in moons_by_parent:
                    body['moons'] = sorted(moons_by_parent[body.get('id')], key=lambda m: float(m.get('orbit', 0)))
                    attach_moons_recursively(body['moons'])
        attach_moons_recursively(system['planets'])
        system['planets'].sort(key=lambda p: (
            0 if any(s in p.get('planet_class', '') for s in ['_star', 'hole', 'pulsar']) and float(p.get('orbit', 0)) == 0 else 2 if any(s in p.get('planet_class', '') for s in ['_star', 'hole', 'pulsar']) else 1,
            float(p.get('orbit', 0))
        ))
        star_count = sum(1 for p in system['planets'] if any(s in p.get('planet_class', '') for s in ['_star', 'hole', 'pulsar']))
        def resolve_names_recursively(body_list, star_context):
            for body in body_list:
                if 'raw_name_block' in body: body['name'] = resolve_name(body['raw_name_block'], loc_data, star_context)
                if 'moons' in body:
                    for moon in body['moons']:
                        if 'raw_name_block' in moon: moon['name'] = resolve_name(moon['raw_name_block'], loc_data, star_context, parent_body_name=body.get('name'))
        resolve_names_recursively(system['planets'], star_count)
        print(f"System {system.get('name', 'Unknown')}: Processed.")
        hierarchical_systems.append(system)
    return hierarchical_systems

def parse_block_content(block_text):
    data = {}
    name_block_content = _get_nested_block_content(block_text, r'name\s*=\s*{')
    if name_block_content: data['raw_name_block'] = name_block_content
    else:
        simple_name_match = re.search(r'^\s*name="([^"]+)"', block_text, re.MULTILINE)
        if simple_name_match: data['name'] = simple_name_match.group(1).replace('_', ' ')
    patterns = {'type': r'^\s*type=([\w_]+)', 'x': r'coordinate=\s*{[^}]*?x=([-\d\.]+)', 'y': r'coordinate=\s*{[^}]*?y=([-\d\.]+)', 'planet_class': r'^\s*planet_class="([^"]+)"', 'planet_size': r'^\s*planet_size=(\d+)', 'orbit': r'^\s*orbit=([-\d\.]+)', 'moon_of': r'^\s*moon_of=(\d+)', 'star_class': r'^\s*star_class="([^"]+)"'}
    for key, pattern in patterns.items():
        match = re.search(pattern, block_text, re.MULTILINE)
        if match: data[key] = match.group(1)
    
    belt_block_content = _get_nested_block_content(block_text, r'asteroid_belts\s*=\s*{')
    if belt_block_content:
        belts_data = []
        type_matches = re.findall(r'type="([^"]+)"', belt_block_content)
        radius_matches = re.findall(r'inner_radius=([-\d\.]+)', belt_block_content)
        
        for i in range(min(len(type_matches), len(radius_matches))):
            belts_data.append({
                'type': type_matches[i],
                'radius': radius_matches[i]
            })
            
        if belts_data:
            data['asteroid_belts_data'] = belts_data

    data['hyperlanes'] = re.findall(r'^\s*to=(\d+)', block_text, re.MULTILINE)
    data['planet_ids'] = re.findall(r'^\s*planet=(\d+)', block_text, re.MULTILINE)
    return data

def parse_nebula_block(block_text):
    """Parses a nebula block for its name, coordinates, and radius."""
    data = {}
    name_block_content = _get_nested_block_content(block_text, r'name\s*=\s*{')
    if name_block_content: data['raw_name_block'] = name_block_content
    patterns = {
        'x': r'coordinate=\s*{[^}]*?x=([-\d\.]+)',
        'y': r'coordinate=\s*{[^}]*?y=([-\d\.]+)',
        'radius': r'^\s*radius=([-\d\.]+)'
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, block_text, re.MULTILINE)
        if match: data[key] = match.group(1)
    return data

def parse_generic_block(block_text):
    """Parses a generic block for simple key=value pairs and coordinates."""
    data = {}
    patterns = {
        'type': r'^\s*type="([^"]+)"',
        'origin': r'^\s*origin=([\d]+)',
        'x': r'coordinate=\s*{[^}]*?x=([-\d\.]+)',
        'y': r'coordinate=\s*{[^}]*?y=([-\d\.]+)',
        'linked_to': r'^\s*linked_to=([\d]+)',
        'bypass': r'^\s*bypass=([\d]+)'
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, block_text, re.MULTILINE)
        if match:
            data[key] = match.group(1)
    return data

def parse_keyed_section(line_iterator, header_regex, block_parser_func):
    """Parses a section of the gamestate with a consistent keyed-block structure."""
    objects = {}
    for line in line_iterator:
        if line.strip() == '}': return objects
        match = header_regex.match(line)
        if match:
            object_id = match.group(1)
            block_lines = [line]
            brace_level = line.count('{') - line.count('}')
            if brace_level <= 0 and '{' in line:
                objects[object_id] = {'id': object_id, **block_parser_func("".join(block_lines))}
                continue
            for block_line in line_iterator:
                block_lines.append(block_line)
                brace_level += block_line.count('{'); brace_level -= block_line.count('}')
                if brace_level <= 0: break
            objects[object_id] = {'id': object_id, **block_parser_func("".join(block_lines))}
    return objects

def parse_stellaris_save(path):
    stars, planets, nebulas, bypasses, natural_wormholes = {}, {}, [], {}, {}
    megastructures_raw = {}
    counts = defaultdict(int)

    try:
        with zipfile.ZipFile(path, 'r') as save_zip:
            if 'gamestate' not in save_zip.namelist(): return None, None, None, None, None, counts
            with save_zip.open('gamestate') as gamestate_file:
                line_iterator = io.TextIOWrapper(gamestate_file, encoding='utf-8')
                star_header_re = re.compile(r'^\t(\d+)=')
                planet_header_re = re.compile(r'^\t\t(\d+)=')
                generic_header_re = re.compile(r'^\t(\d+)=')

                for line in line_iterator:
                    stripped_line = line.strip()
                    if stripped_line == 'galactic_object=': next(line_iterator); stars = parse_keyed_section(line_iterator, star_header_re, parse_block_content)
                    elif stripped_line == 'planets=': next(line_iterator); next(line_iterator); planets = parse_keyed_section(line_iterator, planet_header_re, parse_block_content)
                    elif stripped_line == 'megastructures=': next(line_iterator); megastructures_raw = parse_keyed_section(line_iterator, generic_header_re, parse_generic_block)
                    elif stripped_line == 'bypasses=': next(line_iterator); bypasses = parse_keyed_section(line_iterator, generic_header_re, parse_generic_block)
                    elif stripped_line == 'natural_wormholes=': next(line_iterator); natural_wormholes = parse_keyed_section(line_iterator, generic_header_re, parse_generic_block)
                    elif stripped_line == 'nebula=':
                        block_lines = [line]; brace_level = line.count('{') - line.count('}')
                        if brace_level <= 0 and '{' in line: nebulas.append(parse_nebula_block("".join(block_lines))); continue
                        for block_line in line_iterator:
                            block_lines.append(block_line)
                            brace_level += block_line.count('{'); brace_level -= block_line.count('}')
                            if brace_level <= 0: break
                        nebulas.append(parse_nebula_block("".join(block_lines)))
    except Exception as e:
        print(f"An error occurred during save file parsing: {e}"); return None, None, None, None, None, counts

    # Process wormholes
    bypass_to_system_map = {}
    for nw_data in natural_wormholes.values():
        if 'bypass' in nw_data and 'origin' in nw_data: bypass_to_system_map[nw_data['bypass']] = nw_data['origin']
    
    wormhole_pairs, processed_bypasses = [], set()
    for bypass_id, bypass_data in bypasses.items():
        if bypass_data.get('type') == 'wormhole' and 'linked_to' in bypass_data and bypass_id not in processed_bypasses:
            partner_id = bypass_data['linked_to']
            system1 = bypass_to_system_map.get(bypass_id)
            system2 = bypass_to_system_map.get(partner_id)
            if system1 and system2:
                pair = tuple(sorted((system1, system2)))
                if pair not in wormhole_pairs:
                    wormhole_pairs.append(pair)
            processed_bypasses.add(bypass_id)
            processed_bypasses.add(partner_id)
    
    parsed_megastructures = [m for m in megastructures_raw.values() if 'type' in m and 'origin' in m and m['origin'] != '4294967295']
    
    counts['wormhole_pair'] = len(wormhole_pairs)
    counts['nebula'] = len(nebulas)
    counts['megastructure'] = len(parsed_megastructures)
    for _, planet_data in planets.items():
        p_class = planet_data.get('planet_class', '')
        if any(s in p_class for s in ['_star', 'hole', 'pulsar']): counts['star'] += 1
        elif p_class == "pc_asteroid": counts['asteroid'] += 1
        elif 'moon_of' in planet_data: counts['moon'] += 1
        else: counts['planet'] += 1
    return stars, planets, nebulas, parsed_megastructures, wormhole_pairs, counts

def write_map_file(systems_list, nebulas_list, wormhole_pairs, output_path, loc_data):
    if not systems_list: return

    wormhole_flags_by_system = {}
    for i, pair in enumerate(wormhole_pairs):
        flag = f"continuum_wormhole_{i}"
        wormhole_flags_by_system[pair[0]] = flag
        wormhole_flags_by_system[pair[1]] = flag
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('static_galaxy_scenario = {\n')
        f.write('\tname = "Continuum"\n\tpriority = 200\n\tsupports_shape = elliptical\n\n')
        f.write('\tnum_empires = { min = 1 max = 1 }\n\tnum_empire_default = 1\n\n')
        f.write('\trandom_hyperlanes = no\n\tcore_radius = 0\n\n')
        f.write('\t# --- System Definitions ---\n')
        for system in systems_list:
            sys_id, sys_name = system.get('id'), system.get('name', f"Sys_{system.get('id')}").replace('"', '')
            sys_x, sys_y = system.get('x', '0'), system.get('y', '0')
            initializer_name = f"continuum_system_init_{sys_id}"
            
            flag_string = ""
            if sys_id in wormhole_flags_by_system:
                flag_string = f' effect = {{ set_star_flag = {wormhole_flags_by_system[sys_id]} }}'

            f.write(f'\tsystem = {{ id = "{sys_id}" name = "{sys_name}" position = {{ x = {sys_x} y = {sys_y} }} initializer = {initializer_name}{flag_string} }}\n')

        f.write('\n\t# --- Hyperlane Definitions ---\n')
        processed_lanes, systems_dict = set(), {s['id']: s for s in systems_list}
        for system_id, system_data in systems_dict.items():
            for target_id in system_data.get('hyperlanes', []):
                if target_id in systems_dict:
                    lane_key = tuple(sorted((system_id, target_id)))
                    if lane_key not in processed_lanes:
                        f.write(f'\tadd_hyperlane = {{ from = "{system_id}" to = "{target_id}" }}\n')
                        processed_lanes.add(lane_key)

        if nebulas_list:
            f.write('\n\t# --- Nebula Definitions ---\n')
            for nebula in nebulas_list:
                nebula_name = resolve_name(nebula.get('raw_name_block', ''), loc_data).replace('"', '')
                nebula_x, nebula_y = nebula.get('x', '0'), nebula.get('y', '0')
                nebula_radius = nebula.get('radius', '30')
                f.write(f'\tnebula = {{ name = "{nebula_name}" position = {{ x = {nebula_x} y = {nebula_y} }} radius = {nebula_radius} }}\n')
        
        f.write('}\n')

def write_initializer_file(systems_list, parsed_megastructures, start_system_id, output_path):
    if not systems_list: return
    
    megastructures_by_system = defaultdict(list)
    for mega in parsed_megastructures:
        megastructures_by_system[mega['origin']].append(mega)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        def write_moons_recursively(moons_list, indent_level):
            last_moon_orbit, last_moon_angle = 0.0, 0.0
            tabs = '\t' * indent_level
            
            sorted_moons = []
            for moon in moons_list:
                moon_info = moon.copy()
                if 'x' in moon and 'y' in moon:
                    moon_info['abs_angle'] = math.degrees(math.atan2(-float(moon.get('y', 0)), -float(moon.get('x', 0))))
                else:
                    moon_info['abs_angle'] = 0
                sorted_moons.append(moon_info)
            sorted_moons.sort(key=lambda m: (float(m.get('orbit', 0)), m['abs_angle']))

            for moon in sorted_moons:
                absolute_orbit = float(moon.get("orbit", 10))
                absolute_angle = moon.get("abs_angle", 0)
                relative_orbit = absolute_orbit - last_moon_orbit
                relative_angle = absolute_angle - last_moon_angle
                if relative_angle > 180: relative_angle -= 360
                if relative_angle < -180: relative_angle += 360

                f.write(f'{tabs}moon = {{\n')
                if "name" in moon:
                    moon_name = moon["name"].replace('"', '')
                    f.write(f'{tabs}\tname = "{moon_name}"\n')
                f.write(f'{tabs}\tclass = "{moon.get("planet_class", "pc_barren_cold")}"\n')
                f.write(f'{tabs}\tsize = {moon.get("planet_size", 5)}\n')
                f.write(f'{tabs}\torbit_distance = {relative_orbit:.2f}\n')
                f.write(f'{tabs}\torbit_angle = {round(relative_angle)}\n')
                
                if 'moons' in moon and moon['moons']: write_moons_recursively(moon['moons'], indent_level + 1)
                f.write(f'{tabs}}}\n')
                
                last_moon_orbit = absolute_orbit
                last_moon_angle = absolute_angle

        for system in systems_list:
            sys_id = system.get('id')
            sys_name = system.get('name', f"Sys_{sys_id}").replace('"', '')
            initializer_name = f"continuum_system_init_{sys_id}"
            star_class = system.get('system_star_class', 'sc_g')

            f.write(f"{initializer_name} = {{\n")
            f.write(f'\tname = "{sys_name}"\n\tclass = "{star_class}"\n')
            
            f.write('\tusage = empire_init\n\n' if sys_id == start_system_id else '\tusage = misc_system_init\n\n')
            
            celestial_bodies = []
            if system.get('planets'):
                for p in system.get('planets', []):
                    body_info = {'type': 'planet', 'data': p, 'orbit': float(p.get('orbit', 0))}
                    if 'x' in p and 'y' in p:
                        body_info['angle'] = math.degrees(math.atan2(-float(p.get('y', 0)), -float(p.get('x', 0))))
                    else:
                        body_info['angle'] = 0 
                    celestial_bodies.append(body_info)

            if sys_id in megastructures_by_system:
                for mega in megastructures_by_system[sys_id]:
                    if not ('gateway' in mega.get('type', '') or 'lgate' in mega.get('type', '')):
                        orbit = math.sqrt(float(mega.get('x', 0))**2 + float(mega.get('y', 0))**2)
                        angle = math.degrees(math.atan2(-float(mega.get('y', 0)), -float(mega.get('x', 0))))
                        celestial_bodies.append({'type': 'megastructure', 'data': mega, 'orbit': orbit, 'angle': angle})
            
            if celestial_bodies:
                max_orbit = max((b['orbit'] for b in celestial_bodies), default=0)
                if max_orbit > 590:
                    scale_factor = 590 / max_orbit
                    for body in celestial_bodies: body['orbit'] *= scale_factor
            
            celestial_bodies.sort(key=lambda x: (x['orbit'], x.get('angle', 0)))

            last_absolute_orbit, last_absolute_angle = 0.0, 0.0
            is_first_body = True
            planets_to_write = [b for b in celestial_bodies if b['type'] == 'planet']

            for body in planets_to_write:
                planet = body['data']
                absolute_orbit = body['orbit']
                absolute_angle = body.get('angle', 0)
                
                relative_orbit = absolute_orbit - last_absolute_orbit
                relative_angle = absolute_angle - last_absolute_angle
                if relative_angle > 180: relative_angle -= 360
                if relative_angle < -180: relative_angle += 360

                if is_first_body:
                    relative_angle = 0
                    absolute_angle = 0
                
                f.write(f'\tplanet = {{\n')
                if "name" in planet:
                    planet_name = planet["name"].replace('"', '')
                    if not planet_name.startswith("NEW COLONY"): f.write(f'\t\tname = "{planet_name}"\n')
                p_class = planet.get("planet_class", "pc_barren")
                f.write(f'\t\tclass = "{p_class}"\n\t\tsize = {planet.get("planet_size", 10)}\n')
                f.write(f'\t\torbit_distance = {relative_orbit:.2f}\n')
                f.write(f'\t\torbit_angle = {round(relative_angle)}\n')
                
                if sys_id == start_system_id and is_first_body and not any(s in p_class for s in ['_star', 'hole', 'pulsar']):
                    f.write('\t\thome_planet = yes\n')
                
                is_first_body = False
                
                if 'moons' in planet and planet['moons']: write_moons_recursively(planet['moons'], 2)
                f.write(f'\t}}\n\n')
                
                last_absolute_orbit = absolute_orbit
                last_absolute_angle = absolute_angle
            
            special_objects = [m for m in celestial_bodies if m['type'] == 'megastructure']
            belts_data = system.get('asteroid_belts_data')

            if special_objects or belts_data:
                f.write('\tinit_effect = {\n')
                
                if belts_data:
                    for belt in belts_data:
                        belt_type = belt.get('type', 'rocky_asteroid_belt')
                        belt_radius = belt.get('radius', 95)
                        # --- MODIFICATION START: Corrected effect command and parameter name ---
                        f.write(f'\t\tadd_asteroid_belt = {{ radius = {belt_radius} type = {belt_type} }}\n')
                        # --- MODIFICATION END ---

                if special_objects:
                    for mega_body in special_objects:
                        mega = mega_body['data']
                        mega_type = mega.get('type')
                        orbit_dist = mega_body['orbit']
                        orbit_angle = mega_body['angle']
                        f.write(f'\t\tspawn_megastructure = {{ type = "{mega_type}" orbit_distance = {orbit_dist:.2f} orbit_angle = {orbit_angle:.2f} }}\n')
                
                f.write('\t}\n\n')

            f.write(f"}}\n\n")

def write_on_actions_file(output_path):
    content = """# These should run after the static galaxy has been generated.

on_game_start = {
	events = {
		continuum_wormhole.1 # spawn wormholes based on flags from the parser
	}
}
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def write_events_file(output_path, num_wormhole_pairs):
    event_calls = ""
    if num_wormhole_pairs > 0:
        for i in range(num_wormhole_pairs):
            event_calls += f"\t\tcontinuum_create_wormhole_pair = {{ NUMBER = {i} }}\n"
    
    content = f"""namespace = continuum_wormhole

# spawn non-random wormholes based on flags set by the python parser
event = {{
	id = continuum_wormhole.1
	is_triggered_only = yes
	hide_window = yes
	
	immediate = {{
{event_calls}	}}
}}
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def write_scripted_effects_file(output_path):
    content = """continuum_create_wormhole_pair = {
	# Find the first system in the pair
	random_system = {
		limit = { has_star_flag = continuum_wormhole_$NUMBER$ }
		# Check if it already has a wormhole, just in case. If not, save it.
		if = {
			limit = { has_natural_wormhole = no }
			save_event_target_as = continuum_wormhole_from
		} 
		# If it does have a wormhole, find a nearby system that doesn't.
		else = {
			closest_system = {
				limit = { has_natural_wormhole = no }
				max_steps = 6
				save_event_target_as = continuum_wormhole_from
			}
		}

		# Find the second system in the pair
		random_system = {
			limit = {
				has_star_flag = continuum_wormhole_$NUMBER$
				NOT = { is_same_value = prev } # ensure it's not the same system
			}
			if = {
				limit = { has_natural_wormhole = no }
				save_event_target_as = continuum_wormhole_to
			} else = {
				closest_system = {
					limit = { has_natural_wormhole = no }
					max_steps = 6
					save_event_target_as = continuum_wormhole_to
				}
			}
		}
	}

	# Create and link the wormholes if both ends were found successfully
	if = {
		limit = {
			exists = event_target:continuum_wormhole_from
			event_target:continuum_wormhole_from = { has_natural_wormhole = no }
			exists = event_target:continuum_wormhole_to
			event_target:continuum_wormhole_to = { has_natural_wormhole = no }
		}
		event_target:continuum_wormhole_from = {
			spawn_natural_wormhole = {
				bypass_type = wormhole
				random_pos = yes
			}
		}
		event_target:continuum_wormhole_to = {
			spawn_natural_wormhole = {
				bypass_type = wormhole
				random_pos = yes
			}
			link_wormholes = event_target:continuum_wormhole_from
		}
	}
}
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def main():
    clear_screen()
    print("--- Continuum Galaxy Parser ---")
    
    stellaris_user_dir = find_stellaris_user_dir()
    if not stellaris_user_dir or not os.path.isdir(stellaris_user_dir):
        print("FATAL ERROR: Could not automatically find the Stellaris user documents directory."); input("Press Enter to exit."); return

    stellaris_install_dir = find_stellaris_install_dir()
    if not stellaris_install_dir or not os.path.isdir(stellaris_install_dir):
        print("FATAL ERROR: Could not automatically find the Stellaris game installation directory."); input("Press Enter to exit."); return

    save_game_dir = os.path.join(stellaris_user_dir, 'save games')
    if not os.path.isdir(save_game_dir):
        print(f"FATAL ERROR: Save game directory not found at '{save_game_dir}'"); input("Press Enter to exit."); return
        
    all_saves = []
    for root, dirs, _ in os.walk(save_game_dir):
        for d in dirs:
            save_path = os.path.join(root, d)
            sav_files = [f for f in os.listdir(save_path) if f.endswith('.sav')]
            if sav_files:
                latest_sav_path = max([os.path.join(save_path, f) for f in sav_files], key=os.path.getmtime)
                version, date = get_save_meta_data(latest_sav_path)
                display_name = f"{d}\\{os.path.basename(latest_sav_path)}"
                all_saves.append({'name': display_name, 'path': latest_sav_path, 'version': version, 'date': date})

    if not all_saves: print("No valid save games found."); input("Press Enter to exit."); return
    
    saves_by_version = defaultdict(list);
    for save in all_saves: saves_by_version[save['version']].append(save)

    def version_sort_key(v_str): parts = re.findall(r'(\d+)', v_str); return [int(p) for p in parts]

    print(f"\nSave game location detected:\n{save_game_dir}\n"); print("Please select a save game to parse:")
    
    save_list_for_selection = []
    for version_str in sorted(saves_by_version.keys(), key=version_sort_key):
        print(f"\n- Game Version {version_str} -")
        for save in saves_by_version[version_str]:
            save_list_for_selection.append(save)
            print(f"  [{len(save_list_for_selection)}] {save['date']} {save['name']}")
    
    choice = -1
    while True:
        try:
            choice_str = input(f"\nEnter a Selection or type 'q' to quit: ").lower()
            if choice_str == 'q': print("Exiting parser. Goodbye!"); return
            choice = int(choice_str)
            if 1 <= choice <= len(save_list_for_selection): break
            else: print("Invalid number. Please enter a number from the list.")
        except ValueError: print("Invalid input. Please enter a number or 'q'.")

    selected_save = save_list_for_selection[choice - 1]
    save_file_path, save_version_str = selected_save['path'], selected_save['version']
    
    try:
        clean_save_version = ".".join(re.findall(r'(\d+)', save_version_str)[0:2])
        if float(clean_save_version) < float(SUPPORTED_STELLARIS_VERSION):
            print("\n--- WARNING ---"); print(f"This save is for Stellaris version {save_version_str}, but this parser is designed for version {SUPPORTED_STELLARIS_VERSION}.0+.")
            print("Please open and re-save your game in the latest version of Stellaris for best results.")
            if input("\nDo you want to continue anyway? (y/n): ").lower() != 'y':
                print("Parsing cancelled."); input("Press Enter to exit."); return
    except Exception: print(f"Warning: Could not parse version string '{save_version_str}'. Version check skipped.")

    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()

    # --- Directory Cleanup and Creation ---
    print("\nCleaning up old mod directories...")
    dirs_to_clean = [os.path.join(script_dir, 'map'), os.path.join(script_dir, 'common'), os.path.join(script_dir, 'events')]
    for d in dirs_to_clean:
        if os.path.isdir(d):
            try:
                shutil.rmtree(d)
                print(f"Removed: {os.path.relpath(d, script_dir)}")
            except OSError as e:
                print(f"Error removing directory {d} : {e.strerror}")

    print("Creating new directory structure...")
    output_map_dir = os.path.join(script_dir, "map", "setup_scenarios")
    output_init_dir = os.path.join(script_dir, "common", "solar_system_initializers")
    output_onactions_dir = os.path.join(script_dir, "common", "on_actions")
    output_effects_dir = os.path.join(script_dir, "common", "scripted_effects")
    output_events_dir = os.path.join(script_dir, "events")

    os.makedirs(output_map_dir, exist_ok=True)
    os.makedirs(output_init_dir, exist_ok=True)
    os.makedirs(output_onactions_dir, exist_ok=True)
    os.makedirs(output_effects_dir, exist_ok=True)
    os.makedirs(output_events_dir, exist_ok=True)
    print("Directory structure created successfully.")

    print(f"\nParsing: {selected_save['name']}...")
    
    game_language = get_stellaris_language(stellaris_user_dir)
    localization = load_localization_data(stellaris_install_dir, game_language)
    if not localization: print("FATAL ERROR: No localization data loaded."); input("Press Enter to exit."); return

    parsed_stars, parsed_planets, parsed_nebulas, parsed_megastructures, wormhole_pairs, counts = parse_stellaris_save(save_file_path)
    
    if parsed_stars and parsed_planets:
        galaxy_data = build_galaxy_hierarchy(parsed_stars, parsed_planets, localization)
        system_count = len(galaxy_data)
        start_system_id = None
        for system in galaxy_data:
            if system.get('name', '').lower() == 'sol': 
                start_system_id = system.get('id')
                break
        if not start_system_id and galaxy_data: 
            start_system_id = galaxy_data[0].get('id')
        
        # --- File Generation ---
        output_map_file = os.path.join(output_map_dir, "continuum.txt")
        output_initializer_file = os.path.join(output_init_dir, "continuum_initializers.txt")
        output_onactions_file = os.path.join(output_onactions_dir, "~~~continuum_on_actions.txt")
        output_events_file = os.path.join(output_events_dir, "continuum_wormhole_events.txt")
        output_effects_file = os.path.join(output_effects_dir, "continuum_wormhole_effects.txt")

        write_map_file(galaxy_data, parsed_nebulas, wormhole_pairs, output_map_file, localization)
        write_initializer_file(galaxy_data, parsed_megastructures, start_system_id, output_initializer_file)
        write_on_actions_file(output_onactions_file)
        write_events_file(output_events_file, len(wormhole_pairs)) # Pass the number of pairs
        write_scripted_effects_file(output_effects_file)
        
        print("\n--- PARSING COMPLETE ---")
        print(f"Found {system_count} systems, {counts['nebula']} nebulas, {counts['star']} stars, {counts['planet']} planets, {counts['moon']} moons, and {counts['asteroid']} asteroids.")
        print(f"Found {counts['megastructure']} megastructures and {counts['wormhole_pair']} wormhole pairs.")
        
        print("\nAll required mod files have been generated:")
        print(f"- {os.path.relpath(output_map_file, script_dir)}")
        print(f"- {os.path.relpath(output_initializer_file, script_dir)}")
        print(f"- {os.path.relpath(output_onactions_file, script_dir)}")
        print(f"- {os.path.relpath(output_events_file, script_dir)}")
        print(f"- {os.path.relpath(output_effects_file, script_dir)}")
        print("\nTo load your imported game, you may now select the Continuum galaxy when starting a New Game.")

    else:
        print("Could not parse critical galaxy data from the save file.")

    input("\nPress Enter to exit.")

if __name__ == "__main__":
    main()