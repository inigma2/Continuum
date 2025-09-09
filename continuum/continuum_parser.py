import zipfile
import re
import io
from collections import defaultdict
import math
import os
import sys
import shutil
import json

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
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
                doc_path = winreg.QueryValueEx(key, "Personal")[0]
            doc_path = os.path.expandvars(doc_path)
            stellaris_dir = os.path.join(doc_path, 'Paradox Interactive', 'Stellaris')
            if os.path.isdir(stellaris_dir):
                return stellaris_dir
        except Exception:
            pass

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                doc_path = winreg.QueryValueEx(key, "Personal")[0]
            stellaris_dir = os.path.join(doc_path, 'Paradox Interactive', 'Stellaris')
            if os.path.isdir(stellaris_dir):
                return stellaris_dir
        except Exception:
            print("Warning: Could not query Windows Registry for Documents path. Using standard fallback.")

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
        with open(library_folders_file, 'r', encoding='utf-8') as f:
            for line in f:
                match = re.search(r'"path"\s+"([^"]+)"', line)
                if match: library_paths.append(match.group(1).replace('\\\\', '\\'))
    except Exception as e:
        print(f"Warning: Could not parse Steam library folders file: {e}")
    for path in library_paths:
        stellaris_path = os.path.join(path, 'steamapps', 'common', 'Stellaris')
        if os.path.isdir(stellaris_path): return stellaris_path
    return None

def find_mod_and_game_files(install_dir, user_dir, sub_path):
    paths_to_scan = []
    paths_to_scan.append(os.path.join(install_dir, sub_path))
    paths_to_scan.append(os.path.join(user_dir, 'mod'))
    workshop_path = os.path.abspath(os.path.join(install_dir, '..', '..', 'workshop', 'content', '281990'))
    if os.path.isdir(workshop_path):
        paths_to_scan.append(workshop_path)

    found_files = []
    for path in paths_to_scan:
        if not os.path.isdir(path): continue
        scan_target = os.path.join(path, sub_path) if sub_path not in path else path
        for root, _, files in os.walk(scan_target):
            for file in files:
                if file.endswith('.txt') or file.endswith('.dds'):
                    found_files.append(os.path.join(root, file))
    return found_files

def _get_nested_block_content(text, start_regex, start_index=0):
    match = re.search(start_regex, text[start_index:])
    if not match: return None, -1, -1

    search_start = start_index + match.start()
    
    try:
        content_start_index = text.index('{', search_start) + 1
    except ValueError:
        return None, -1, -1

    brace_level = 1
    for i in range(content_start_index, len(text)):
        char = text[i]
        if char == '{': brace_level += 1
        elif char == '}': brace_level -= 1
        if brace_level == 0:
            return text[content_start_index:i], search_start, i + 1
    return None, -1, -1

def parse_all_megastructures(file_list):
    definitions = {}
    var_pattern = re.compile(r"^\s*@([\w_]+)\s*=\s*([-\d.]+)", re.MULTILINE)
    
    for file_path in file_list:
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                content = f.read()
                
                local_vars = {f"@{m.group(1)}": m.group(2) for m in var_pattern.finditer(content)}
                
                brace_level = 0
                in_comment = False
                current_key = None
                block_start = 0
                
                for i, char in enumerate(content):
                    if char == '#': in_comment = True
                    elif char == '\n': in_comment = False
                    if in_comment: continue

                    if char == '{':
                        if brace_level == 0:
                            key_match = re.search(r'([\w_]+)\s*=\s*$', content[:i])
                            if key_match:
                                current_key = key_match.group(1)
                                block_start = i + 1
                        brace_level += 1
                    elif char == '}':
                        brace_level -= 1
                        if brace_level == 0 and current_key:
                            block_content = content[block_start:i]

                            for var, val in local_vars.items():
                                block_content = block_content.replace(var, val)

                            star_flags = []
                            country_flags = []
                            
                            on_complete_content, _, _ = _get_nested_block_content(block_content, r'on_build_complete\s*=\s*{')
                            if on_complete_content:
                                star_flags.extend(re.findall(r'set_star_flag\s*=\s*([\w_]+)', on_complete_content))
                                
                                from_block_content, _, _ = _get_nested_block_content(on_complete_content, r'(?:from|owner)\s*=\s*{')
                                if from_block_content:
                                    country_flags.extend(re.findall(r'set_country_flag\s*=\s*([\w_]+)', from_block_content))

                            definitions[current_key] = {
                                'content': block_content,
                                'star_flags': list(set(star_flags)), 
                                'country_flags': list(set(country_flags))
                            }
                            current_key = None
        except Exception as e:
            print(f"Warning: Could not read or parse {file_path}: {e}")
    print(f"Parsed {len(definitions)} megastructure definitions from game and mod files.")
    return definitions

def get_save_meta_data(save_file_path):
    version = "Unknown"; date = "Unknown Date"
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
    print(f"\nSearching for all localization files in:\n{base_loc_path}\n")
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
    
    if 'key="HABITAT_PLANET_NAME"' in name_block_content and 'key="FROM.from.solar_system.GetName"' in name_block_content:
        system_name_match = re.search(r'key="FROM\.from\.solar_system\.GetName"\s*value\s*=\s*{\s*key="([^"]+)"', name_block_content)
        if system_name_match:
            system_name = system_name_match.group(1)
            resolved_system_name = loc_data.get(system_name, system_name.replace('_', ' '))
            return f"{resolved_system_name} Habitat Complex"

    key_match = re.search(r'^\s*key="([^"]+)"', name_block_content)
    if not key_match:
        key_match_simple = re.search(r'key="([^"]+)"', name_block_content)
        if not key_match_simple: return "Unknown"
        name_key = key_match_simple.group(1)
    else:
        name_key = key_match.group(1)

    if name_key.startswith('$') and name_key.endswith('$'): name_key = name_key.strip('$')
    variables_content,_,_ = _get_nested_block_content(name_block_content, r'variables\s*=\s*{')
    if (name_key.startswith("STAR_NAME_") or name_key.endswith("_NAME_FORMAT") or name_key.startswith("NEW_COLONY_NAME")) and variables_content:
        if name_key.startswith("STAR_NAME_") or name_key.startswith("NEW_COLONY_NAME"):
            name_value_block,_,_ = _get_nested_block_content(variables_content, r'key="NAME"\s*value\s*=\s*{')
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
            parent_value_block,_,_ = _get_nested_block_content(variables_content, r'key="PARENT"\s*value\s*=\s*{')
            numeral_value_block,_,_ = _get_nested_block_content(variables_content, r'key="NUMERAL"\s*value\s*=\s*{')
            if parent_value_block and numeral_value_block:
                parent_name_val = resolve_name(parent_value_block, loc_data, star_count_context)
                numeral_key_match = re.search(r'key="([^"]+)"', numeral_value_block)
                if numeral_key_match: return f"{parent_name_val} {numeral_key_match.group(1)}"
            return "Unknown Planet"
        if name_key == "SUBPLANET_NAME_FORMAT":
            parent_value_block,_,_ = _get_nested_block_content(variables_content, r'key="PARENT"\s*value\s*=\s*{')
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
            prefix_val_block,_,_ = _get_nested_block_content(variables_content, r'key="prefix"\s*value\s*=\s*{')
            if prefix_val_block:
                prefix_match = re.search(r'key="([^"]+)"', prefix_val_block)
                if prefix_match: prefix = prefix_match.group(1)
            suffix_val_block,_,_ = _get_nested_block_content(variables_content, r'key="suffix"\s*value\s*=\s*{')
            if suffix_val_block:
                suffix_match = re.search(r'key="([^"]+)"', suffix_val_block)
                if suffix_match: suffix = suffix_match.group(1)
            return f"{prefix}{suffix}"
    if name_key in loc_data: return loc_data[name_key]
    clean_name = re.sub(r'(_system|_SYSTEM)$', '', name_key)
    clean_name = re.sub(r'^(NAME_|SPEC_)', '', clean_name)
    return clean_name.replace('_', ' ')

def build_galaxy_hierarchy(stars, planets, loc_data):
    """
    Builds a detailed hierarchical map of each star system based on explicit save game structure.
    """
    hierarchical_systems = []
    for star_id, star_data in stars.items():
        system = star_data
        system['system_star_class'] = system.get('star_class', 'sc_g')
        if 'raw_name_block' in star_data:
            system['name'] = resolve_name(star_data['raw_name_block'], loc_data)
        
        all_bodies_in_system_map = {}
        for p_id in star_data.get('planet_ids', []):
            if p_id in planets:
                body = planets[p_id]
                body['abs_x'] = float(body.get('x', '0'))
                body['abs_y'] = float(body.get('y', '0'))
                body['children'] = []
                body['parent'] = None
                body['body_type'] = 'star' if any(s in body.get('planet_class', '') for s in ['_star', 'hole', 'pulsar']) else 'planet'
                all_bodies_in_system_map[p_id] = body

        system_center = {'id': '0', 'abs_x': 0.0, 'abs_y': 0.0, 'children': [], 'nesting_level': 0, 'name': 'System Center'}
        
        # Explicitly parent moons first, based on 'moon_of' tag
        for body_id, body in all_bodies_in_system_map.items():
            if 'moon_of' in body:
                parent_id = body['moon_of']
                if parent_id in all_bodies_in_system_map:
                    parent_body = all_bodies_in_system_map[parent_id]
                    parent_body['children'].append(body)
                    body['parent'] = parent_body
        
        # Parent all remaining unparented bodies to the system center
        for body_id, body in all_bodies_in_system_map.items():
            if body['parent'] is None:
                system_center['children'].append(body)
                body['parent'] = system_center
        
        def set_nesting_levels(body, level):
            body['nesting_level'] = level
            for child in body.get('children', []):
                set_nesting_levels(child, level + 1)
        
        set_nesting_levels(system_center, 0)
        
        system['hierarchy_root'] = system_center

        star_count = len([b for b in all_bodies_in_system_map.values() if b['body_type'] == 'star'])
        def resolve_all_names(body):
            if 'raw_name_block' in body:
                parent_name = body.get('parent', {}).get('name')
                body['name'] = resolve_name(body['raw_name_block'], loc_data, star_count, parent_body_name=parent_name)
            for child in body.get('children', []):
                resolve_all_names(child)
        
        resolve_all_names(system_center)

        print(f"System {system.get('name', 'Unknown')}: Processed hierarchy.")
        hierarchical_systems.append(system)
    
    return hierarchical_systems

def parse_block_content(block_text):
    data = {}
    name_block_content,_,_ = _get_nested_block_content(block_text, r'name\s*=\s*{')
    if name_block_content: data['raw_name_block'] = name_block_content
    else:
        simple_name_match = re.search(r'^\s*name="([^"]+)"', block_text, re.MULTILINE)
        if simple_name_match: data['name'] = simple_name_match.group(1).replace('_', ' ')
    patterns = {'type': r'^\s*type=([\w_]+)', 'x': r'coordinate=\s*{[^}]*?x=([-\d\.]+)', 'y': r'coordinate=\s*{[^}]*?y=([-\d\.]+)', 'planet_class': r'^\s*planet_class="([^"]+)"', 'planet_size': r'^\s*planet_size=(\d+)', 'orbit': r'^\s*orbit=([-\d\.]+)', 'moon_of': r'^\s*moon_of=(\d+)', 'star_class': r'^\s*star_class="([^"]+)"'}
    for key, pattern in patterns.items():
        match = re.search(pattern, block_text, re.MULTILINE)
        if match: data[key] = match.group(1)
    
    belt_block_content,_,_ = _get_nested_block_content(block_text, r'asteroid_belts\s*=\s*{')
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
    data = {}
    name_block_content,_,_ = _get_nested_block_content(block_text, r'name\s*=\s*{')
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
    data = {}
    name_match = re.search(r'^\s*name="([^"]+)"', block_text, re.MULTILINE)
    if name_match:
        data['name'] = name_match.group(1)
    
    patterns = {
        'type': r'^\s*type="([^"]+)"',
        'origin': r'coordinate=\s*{[^}]*?origin=([\d\.]+)',
        'x': r'coordinate=\s*{[^}]*?x=([-\d\.]+)',
        'y': r'coordinate=\s*{[^}]*?y=([-\d\.]+)',
        'linked_to': r'^\s*linked_to=([\d]+)',
        'bypass': r'^\s*bypass=([\d]+)',
        'graphical_culture': r'^\s*graphical_culture="([^"]+)"',
        'owner': r'^\s*owner=([\d]+)',
        'planet': r'^\s*planet=([\d]+)'
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, block_text, re.MULTILINE)
        if match:
            data[key] = match.group(1)
    return data

def parse_keyed_section(line_iterator, header_regex, block_parser_func):
    objects = {}
    for line in line_iterator:
        if line.strip() == '}': return objects
        match = header_regex.match(line)
        if match:
            if "none" in line:
                continue
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
            processed_bypasses.add(bypass_id); processed_bypasses.add(partner_id)
    
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

def write_initializer_file(systems_list, parsed_megastructures, start_system_id, output_path, all_mega_definitions):
    if not systems_list: return
    
    megastructures_by_system = defaultdict(list)
    for mega in parsed_megastructures:
        original_planet_id = mega.get('planet')
        if not original_planet_id or original_planet_id == '4294967295':
            megastructures_by_system[mega['origin']].append(mega)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        # --- START OF 7 ORBIT CALCULATION FUNCTIONS (LOGICAL GROUPING) ---

        def _calculate_orbit_params(body, parent):
            """Core calculation logic based on absolute coordinates."""
            if not parent:
                print(f"ERROR: Body '{body.get('name', body.get('id'))}' has a None parent. Skipping calculation.")
                return {'distance': 0, 'angle': 0}
            rel_x = body['abs_x'] - parent['abs_x']
            rel_y = body['abs_y'] - parent['abs_y']
            distance = math.sqrt(rel_x**2 + rel_y**2)
            angle = math.degrees(math.atan2(-rel_y, -rel_x))
            return {'distance': distance, 'angle': angle}

        def calculate_level_1_orbit(body, parent):
            """Nesting Level 1: Bodies orbiting the system center (barycenter)."""
            return _calculate_orbit_params(body, parent)

        def calculate_level_2_orbit(body, parent):
            """Nesting Level 2: Bodies orbiting a Level 1 body."""
            return _calculate_orbit_params(body, parent)

        def calculate_level_3_orbit(body, parent):
            """Nesting Level 3: Bodies orbiting a Level 2 body."""
            return _calculate_orbit_params(body, parent)
        
        def calculate_level_4_orbit(body, parent):
            """Nesting Level 4: Bodies orbiting a Level 3 body."""
            return _calculate_orbit_params(body, parent)

        def calculate_deeper_level_orbit(body, parent):
            """Nesting Level 5+: Bodies in deeper orbits."""
            return _calculate_orbit_params(body, parent)
        
        # --- END OF 7 ORBIT CALCULATION FUNCTIONS ---

        def write_body_init_effects(body, tabs):
            init_effects = []
            if 'attached_mega' in body:
                mega = body['attached_mega']
                mega_type = mega.get("type")
                mega_gfx = mega.get("graphical_culture", "none")
                flag_name = f"continuum_host_{mega_type}_{mega_gfx}"
                init_effects.append(f'{tabs}\tset_planet_flag = {flag_name}')
                if 'name' in mega:
                    clean_name = mega["name"].replace('"', '\\"')
                    init_effects.append(f'{tabs}\tset_variable = {{ which = continuum_mega_name value = "{clean_name}" }}')
            if body.get("planet_class") == "pc_habitat":
                init_effects.append(f'{tabs}\tset_planet_flag = habitat')
            
            if init_effects:
                f.write(f'{tabs}init_effect = {{\n')
                f.write('\n'.join(init_effects) + '\n')
                f.write(f'{tabs}}}\n')

        for system in systems_list:
            sys_id = system.get('id')
            sys_name = system.get('name', f"Sys_{sys_id}").replace('"', '')
            initializer_name = f"continuum_system_init_{sys_id}"
            star_class = system.get('system_star_class', 'sc_g')

            f.write(f"{initializer_name} = {{\n")
            f.write(f'\tname = "{sys_name}"\n\tclass = "{star_class}"\n')
            f.write('\tusage = empire_init\n\n' if sys_id == start_system_id else '\tusage = misc_system_init\n\n')

            # --- START OF NON-RECURSIVE ORCHESTRATOR ---
            
            hierarchy_root = system.get('hierarchy_root')
            if not hierarchy_root:
                f.write("}\n\n")
                continue

            # --- Automatic System Scaling ---
            max_radius = 0
            all_bodies_in_system = []
            queue = [hierarchy_root]
            while queue:
                body = queue.pop(0)
                all_bodies_in_system.append(body)
                queue.extend(body.get('children', []))

            for body in all_bodies_in_system[1:]: # Skip root
                radius = math.sqrt(body['abs_x']**2 + body['abs_y']**2)
                if radius > max_radius:
                    max_radius = radius
            
            scale_factor = 1.0
            if max_radius > 590:
                scale_factor = 590 / max_radius
                print(f"INFO: System '{sys_name}' is too large (radius: {max_radius:.2f}). Scaling by {scale_factor:.2f}.")
                for body in all_bodies_in_system:
                    body['abs_x'] *= scale_factor
                    body['abs_y'] *= scale_factor

            # --- True Barycentric Writing ---
            level_1_bodies = sorted(hierarchy_root['children'], key=lambda b: _calculate_orbit_params(b, hierarchy_root)['distance'])
            
            last_orbit_l1, last_angle_l1 = 0.0, 0.0 
            is_first_l1_body = True
            for body_l1 in level_1_bodies:
                orbit_params_l1 = calculate_level_1_orbit(body_l1, body_l1['parent'])
                rel_dist_l1 = orbit_params_l1['distance'] - last_orbit_l1
                rel_angle_l1 = orbit_params_l1['angle'] - last_angle_l1
                if rel_angle_l1 > 180: rel_angle_l1 -= 360
                if rel_angle_l1 < -180: rel_angle_l1 += 360

                f.write(f'\tplanet = {{\n')
                if "name" in body_l1:
                    clean_name = body_l1["name"].replace('"', '')
                    f.write(f'\t\tname = "{clean_name}"\n')
                f.write(f'\t\tclass = "{body_l1.get("planet_class", "pc_barren")}"\n\t\tsize = {body_l1.get("planet_size", 10)}\n')
                
                if is_first_l1_body:
                    f.write(f'\t\torbit_distance = {orbit_params_l1["distance"]:.2f}\n')
                    f.write(f'\t\torbit_angle = {round(orbit_params_l1["angle"])}\n')
                    is_first_l1_body = False
                else:
                    f.write(f'\t\torbit_distance = {rel_dist_l1:.2f}\n\t\torbit_angle = {round(rel_angle_l1)}\n')

                if sys_id == start_system_id and body_l1 is level_1_bodies[0] and body_l1['body_type'] != 'star':
                     f.write('\t\thome_planet = yes\n')
                
                write_body_init_effects(body_l1, '\t\t')
                
                last_orbit_l2, last_angle_l2 = 0.0, 0.0
                children_l2 = sorted(body_l1.get('children', []), key=lambda b: _calculate_orbit_params(b, body_l1)['distance'])
                is_first_l2_body = True
                for body_l2 in children_l2:
                    orbit_params_l2 = calculate_level_2_orbit(body_l2, body_l2['parent'])
                    rel_dist_l2 = orbit_params_l2['distance'] - last_orbit_l2
                    rel_angle_l2 = orbit_params_l2['angle'] - last_angle_l2
                    if rel_angle_l2 > 180: rel_angle_l2 -= 360
                    if rel_angle_l2 < -180: rel_angle_l2 += 360
                    
                    f.write(f'\t\tmoon = {{\n')
                    if "name" in body_l2:
                        clean_name = body_l2["name"].replace('"', '')
                        f.write(f'\t\t\tname = "{clean_name}"\n')
                    f.write(f'\t\t\tclass = "{body_l2.get("planet_class", "pc_barren")}"\n\t\t\tsize = {body_l2.get("planet_size", 10)}\n')
                    
                    if is_first_l2_body:
                        f.write(f'\t\t\torbit_distance = {orbit_params_l2["distance"]:.2f}\n')
                        f.write(f'\t\t\torbit_angle = {round(orbit_params_l2["angle"])}\n')
                        is_first_l2_body = False
                    else:
                        f.write(f'\t\t\torbit_distance = {rel_dist_l2:.2f}\n\t\t\torbit_angle = {round(rel_angle_l2)}\n')
                    
                    write_body_init_effects(body_l2, '\t\t\t')

                    last_orbit_l3, last_angle_l3 = 0.0, 0.0
                    children_l3 = sorted(body_l2.get('children', []), key=lambda b: _calculate_orbit_params(b, body_l2)['distance'])
                    is_first_l3_body = True
                    for body_l3 in children_l3:
                        orbit_params_l3 = calculate_level_3_orbit(body_l3, body_l3['parent'])
                        rel_dist_l3 = orbit_params_l3['distance'] - last_orbit_l3
                        rel_angle_l3 = orbit_params_l3['angle'] - last_angle_l3
                        if rel_angle_l3 > 180: rel_angle_l3 -= 360
                        if rel_angle_l3 < -180: rel_angle_l3 += 360

                        f.write(f'\t\t\tmoon = {{\n')
                        if "name" in body_l3:
                            clean_name = body_l3["name"].replace('"', '')
                            f.write(f'\t\t\t\tname = "{clean_name}"\n')
                        f.write(f'\t\t\t\tclass = "{body_l3.get("planet_class", "pc_barren")}"\n\t\t\tsize = {body_l3.get("planet_size", 10)}\n')
                        
                        if is_first_l3_body:
                             f.write(f'\t\t\t\torbit_distance = {orbit_params_l3["distance"]:.2f}\n')
                             f.write(f'\t\t\t\torbit_angle = {round(orbit_params_l3["angle"])}\n')
                             is_first_l3_body = False
                        else:
                            f.write(f'\t\t\t\torbit_distance = {rel_dist_l3:.2f}\n\t\t\t\torbit_angle = {round(rel_angle_l3)}\n')

                        write_body_init_effects(body_l3, '\t\t\t\t')
                        f.write(f'\t\t\t}}\n')
                        last_orbit_l3, last_angle_l3 = orbit_params_l3['distance'], orbit_params_l3['angle']
                    
                    f.write(f'\t\t}}\n')
                    last_orbit_l2, last_angle_l2 = orbit_params_l2['distance'], orbit_params_l2['angle']

                f.write('\t}\n\n')
                last_orbit_l1, last_angle_l1 = orbit_params_l1['distance'], orbit_params_l1['angle']

            # --- Add init_effect for belts and space-born megastructures ---
            has_belts = system.get('asteroid_belts_data')
            has_megas = sys_id in megastructures_by_system
            if has_belts or has_megas:
                f.write('\tinit_effect = {\n')
                if has_belts:
                    for belt in system.get('asteroid_belts_data'):
                        belt_type = belt.get('type', 'rocky_asteroid_belt')
                        belt_radius = float(belt.get('radius', 95)) * scale_factor
                        f.write(f'\t\tadd_asteroid_belt = {{ radius = {belt_radius:.2f} type = {belt_type} }}\n')
                if has_megas:
                    for mega in megastructures_by_system[sys_id]:
                        mega_type = mega.get("type")
                        param_dict = {'type': f'type = {mega_type}'}
                        if 'name' in mega:
                            clean_name = mega["name"].replace('"', '\\"')
                            param_dict['name'] = f'name = "{clean_name}"'
                        if 'graphical_culture' in mega:
                             param_dict['graphical_culture'] = f'graphical_culture = {mega["graphical_culture"]}'
                        
                        # Use scaled coordinates for megastructure placement
                        mega_x = float(mega.get('x', '0')) * scale_factor
                        mega_y = float(mega.get('y', '0')) * scale_factor
                        param_dict['orbit_distance'] = f'orbit_distance = {math.sqrt(mega_x**2 + mega_y**2):.2f}'
                        param_dict['orbit_angle'] = f'orbit_angle = {math.degrees(math.atan2(-mega_y, -mega_x)):.2f}'
                        
                        order = ['type', 'name', 'graphical_culture', 'orbit_distance', 'orbit_angle']
                        params = [param_dict[key] for key in order if key in param_dict]
                        param_string = " ".join(params)
                        f.write(f'\t\tspawn_megastructure = {{ {param_string} }}\n')

                        mega_def = all_mega_definitions.get(mega_type, {})
                        star_flags = mega_def.get('star_flags', [])
                        for flag in star_flags:
                            f.write(f'\t\tset_star_flag = {flag}\n')
                f.write('\t}\n')
            
            f.write(f"}}\n\n")

def find_body_in_system(hierarchy_root, target_id):
    if not hierarchy_root: return None
    queue = [hierarchy_root]
    while queue:
        body = queue.pop(0)
        if body.get('id') == target_id:
            return body
        if 'children' in body:
            queue.extend(body['children'])
    return None

def write_on_actions_file(output_path, has_wormholes, has_planet_megas):
    content = "# These should run after the static galaxy has been generated.\n\non_game_start = {\n\tevents = {\n"
    if has_wormholes:
        content += "\t\tcontinuum_wormhole.1 # spawn wormholes based on flags from the parser\n"
    if has_planet_megas:
        content += "\t\tcontinuum_megastructure.1 # spawn planet-bound megastructures\n"
    content += "\t}\n}\n"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def write_wormhole_events_file(output_path, num_wormhole_pairs):
    if num_wormhole_pairs == 0: return
    event_calls = ""
    for i in range(num_wormhole_pairs):
        event_calls += f"\t\tcontinuum_create_wormhole_pair = {{ NUMBER = {i} }}\n"
    
    content = f"""namespace = continuum_wormhole
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

def write_megastructure_events_file(output_path, planet_megas):
    """Generates the event file based on verified commands and syntax."""
    if not planet_megas: return

    unique_megas = set()
    for mega in planet_megas:
        mega_type = mega.get("type")
        mega_gfx = mega.get("graphical_culture", "none")
        unique_megas.add((mega_type, mega_gfx))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("namespace = continuum_megastructure\n\n")
        f.write("event = {\n")
        f.write("\tid = continuum_megastructure.1\n")
        f.write("\tis_triggered_only = yes\n")
        f.write("\thide_window = yes\n\n")
        f.write("\timmediate = {\n")
        f.write("\t\tevery_galaxy_planet = {\n")
        
        f.write("\t\t\tlimit = { OR = { ")
        for mega_type, mega_gfx in sorted(list(unique_megas)):
            flag_name = f"continuum_host_{mega_type}_{mega_gfx}"
            f.write(f"has_planet_flag = {flag_name} ")
        f.write("} }\n\n")
        
        is_first = True
        for mega_type, mega_gfx in sorted(list(unique_megas)):
            flag_name = f"continuum_host_{mega_type}_{mega_gfx}"
            
            if_statement = "if" if is_first else "else_if"
            is_first = False
            
            f.write(f"\t\t\t{if_statement} = {{\n")
            f.write(f"\t\t\t\tlimit = {{ has_planet_flag = {flag_name} }}\n\n")
            
            f.write("\t\t\t\tsolar_system = {\n")
            f.write("\t\t\t\t\tif = {\n")
            f.write("\t\t\t\t\t\tlimit = { prev = { is_variable_set = continuum_mega_name } }\n")
            f.write("\t\t\t\t\t\tspawn_megastructure = {\n")
            f.write(f"\t\t\t\t\t\t\ttype = {mega_type}\n")
            f.write("\t\t\t\t\t\t\tplanet = prev\n")
            if mega_gfx != "none":
                f.write(f'\t\t\t\t\t\t\tgraphical_culture = {mega_gfx}\n')
            f.write('\t\t\t\t\t\t\tname = "[prev.continuum_mega_name]"\n')
            f.write("\t\t\t\t\t\t}\n")
            f.write("\t\t\t\t\t}\n")
            f.write("\t\t\t\t\telse = {\n")
            f.write("\t\t\t\t\t\tspawn_megastructure = {\n")
            f.write(f"\t\t\t\t\t\t\ttype = {mega_type}\n")
            f.write("\t\t\t\t\t\t\tplanet = prev\n")
            if mega_gfx != "none":
                f.write(f'\t\t\t\t\t\t\tgraphical_culture = {mega_gfx}\n')
            f.write("\t\t\t\t\t\t}\n")
            f.write("\t\t\t\t\t}\n")
            f.write("\t\t\t\t}\n\n")
            
            f.write(f"\t\t\t\tremove_planet_flag = {flag_name}\n")
            f.write("\t\t\t\tclear_variable = continuum_mega_name\n")
            f.write("\t\t\t}\n")
        
        f.write("\t\t}\n")
        f.write("\t}\n")
        f.write("}\n")

def write_scripted_effects_file(output_path, num_wormhole_pairs):
    content = ""
    if num_wormhole_pairs > 0:
        content += """continuum_create_wormhole_pair = {
	random_system = {
		limit = { has_star_flag = continuum_wormhole_$NUMBER$ }
		if = {
			limit = { has_natural_wormhole = no }
			save_event_target_as = continuum_wormhole_from
		} 
		else = {
			closest_system = {
				limit = { has_natural_wormhole = no }
				max_steps = 6
				save_event_target_as = continuum_wormhole_from
			}
		}
		random_system = {
			limit = {
				has_star_flag = continuum_wormhole_$NUMBER$
				NOT = { is_same_value = prev }
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
	if = {
		limit = {
			exists = event_target:continuum_wormhole_from
			event_target:continuum_wormhole_from = { has_natural_wormhole = no }
			exists = event_target:continuum_wormhole_to
			event_target:continuum_wormhole_to = { has_natural_wormhole = no }
		}
		event_target:continuum_wormhole_from = { spawn_natural_wormhole = { bypass_type = wormhole random_pos = yes } }
		event_target:continuum_wormhole_to = { spawn_natural_wormhole = { bypass_type = wormhole random_pos = yes } link_wormholes = event_target:continuum_wormhole_from }
	}
}
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

def main():
    clear_screen()
    print("--- Continuum Galaxy Parser ---")
    
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()

    debug_file_path = os.path.join(script_dir, 'parserdebug.txt')
    if os.path.exists(debug_file_path):
        os.remove(debug_file_path)

    stellaris_user_dir = find_stellaris_user_dir()
    if not stellaris_user_dir or not os.path.isdir(stellaris_user_dir):
        print("FATAL ERROR: Could not find Stellaris user directory."); input("Press Enter to exit."); return

    stellaris_install_dir = find_stellaris_install_dir()
    if not stellaris_install_dir or not os.path.isdir(stellaris_install_dir):
        print("FATAL ERROR: Could not find Stellaris install directory."); input("Press Enter to exit."); return

    save_game_dir = os.path.join(stellaris_user_dir, 'save games')
    if not os.path.isdir(save_game_dir):
        print(f"FATAL ERROR: Save game directory not found at '{save_game_dir}'"); input("Press Enter to exit."); return
    
    all_saves = []
    for root, _, files in os.walk(save_game_dir):
        relative_dir_path = os.path.relpath(root, save_game_dir)
        for file in files:
            if file.endswith('.sav'):
                full_sav_path = os.path.join(root, file)
                version, date = get_save_meta_data(full_sav_path)
                display_name = os.path.join(relative_dir_path, file) if relative_dir_path != '.' else file
                all_saves.append({'name': display_name, 'path': full_sav_path, 'version': version, 'date': date})

    if not all_saves: print("No valid save games found."); input("Press Enter to exit."); return
    
    saves_by_version = defaultdict(list);
    for save in all_saves: saves_by_version[save['version']].append(save)

    def version_sort_key(v_str): parts = re.findall(r'(\d+)', v_str); return [int(p) for p in parts]

    print(f"\nSave game location: {save_game_dir}\n"); print("Please select a save game:")
    
    save_list_for_selection = []
    for version_str in sorted(saves_by_version.keys(), key=version_sort_key):
        print(f"\n- Game Version {version_str} -")
        sorted_saves = sorted(saves_by_version[version_str], key=lambda x: x['date'], reverse=True)
        for save in sorted_saves:
            save_list_for_selection.append(save)
            print(f"  [{len(save_list_for_selection)}] {save['date']} {save['name']}")
    
    choice = -1
    while True:
        try:
            choice_str = input(f"\nEnter a Selection or 'q' to quit: ").lower()
            if choice_str == 'q': print("Exiting."); return
            choice = int(choice_str)
            if 1 <= choice <= len(save_list_for_selection): break
            else: print("Invalid number.")
        except ValueError: print("Invalid input.")

    selected_save = save_list_for_selection[choice - 1]
    save_file_path, save_version_str = selected_save['path'], selected_save['version']
    
    try:
        clean_save_version = ".".join(re.findall(r'(\d+)', save_version_str)[0:2])
        if float(clean_save_version) < float(SUPPORTED_STELLARIS_VERSION):
            print(f"\n--- WARNING: This save is for Stellaris {save_version_str}, but this parser is for {SUPPORTED_STELLARIS_VERSION}.0+. ---")
            print("Please re-save your game in the latest version of Stellaris for best results.")
            if input("Continue anyway? (y/n): ").lower() != 'y':
                print("Parsing cancelled."); input("Press Enter to exit."); return
    except Exception: print(f"Warning: Could not parse version string '{save_version_str}'.")

    print("\nCleaning up old mod directories...")
    dirs_to_clean = [os.path.join(script_dir, d) for d in ['map', 'common', 'events']]
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

    for d in [output_map_dir, output_init_dir, output_onactions_dir, output_effects_dir, output_events_dir]:
        os.makedirs(d, exist_ok=True)
    print("Directory structure created.")

    print(f"\nParsing: {selected_save['name']}...")
    
    print("\nScanning for megastructure definitions...")
    mega_files = find_mod_and_game_files(stellaris_install_dir, stellaris_user_dir, 'common/megastructures')
    all_mega_definitions = parse_all_megastructures(mega_files)

    game_language = get_stellaris_language(stellaris_user_dir)
    localization = load_localization_data(stellaris_install_dir, game_language)
    if not localization: print("FATAL ERROR: No localization data loaded."); input("Press Enter to exit."); return

    parsed_stars, parsed_planets, parsed_nebulas, parsed_megastructures, wormhole_pairs, counts = parse_stellaris_save(save_file_path)
    
    if parsed_stars and parsed_planets:
        galaxy_data = build_galaxy_hierarchy(parsed_stars, parsed_planets, localization)
        
        print("Attaching planet-bound megastructures to hosts...")
        planet_bound_megas = []
        systems_map = {sys['id']: sys for sys in galaxy_data}

        for mega in parsed_megastructures:
            original_planet_id = mega.get('planet')
            if original_planet_id and original_planet_id != '4294967295' and mega.get('origin') in systems_map:
                planet_bound_megas.append(mega)
                target_system = systems_map[mega['origin']]
                host_planet = find_body_in_system(target_system.get('hierarchy_root'), original_planet_id)
                if host_planet:
                    host_planet['attached_mega'] = mega
                else:
                    print(f"Warning: Could not find host planet ID {original_planet_id} for megastructure '{mega.get('type')}' in system {target_system.get('name')}")
        
        system_count = len(galaxy_data)
        start_system_id = None
        for system in galaxy_data:
            if system.get('name', '').lower() == 'sol': 
                start_system_id = system.get('id')
                break
        if not start_system_id and galaxy_data: 
            start_system_id = galaxy_data[0].get('id')
        
        output_map_file = os.path.join(output_map_dir, "continuum.txt")
        output_initializer_file = os.path.join(output_init_dir, "continuum_initializers.txt")
        output_onactions_file = os.path.join(output_onactions_dir, "~~~continuum_on_actions.txt")
        output_wormhole_events_file = os.path.join(output_events_dir, "continuum_wormhole_events.txt")
        output_wormhole_effects_file = os.path.join(output_effects_dir, "continuum_wormhole_effects.txt")
        output_mega_events_file = os.path.join(output_events_dir, "continuum_megastructure_events.txt")

        write_map_file(galaxy_data, parsed_nebulas, wormhole_pairs, output_map_file, localization)
        write_initializer_file(galaxy_data, parsed_megastructures, start_system_id, output_initializer_file, all_mega_definitions)
        write_on_actions_file(output_onactions_file, has_wormholes=(len(wormhole_pairs) > 0), has_planet_megas=(len(planet_bound_megas) > 0))
        write_wormhole_events_file(output_wormhole_events_file, len(wormhole_pairs))
        write_megastructure_events_file(output_mega_events_file, planet_bound_megas)
        write_scripted_effects_file(output_wormhole_effects_file, len(wormhole_pairs))
        
        print("\n--- PARSING COMPLETE ---")
        print(f"Found {system_count} systems, {counts['nebula']} nebulas, {counts['star']} stars, {counts['planet']} planets, {counts['moon']} moons, and {counts['asteroid']} asteroids.")
        print(f"Found {counts['megastructure']} megastructures ({len(planet_bound_megas)} planet-bound) and {counts['wormhole_pair']} wormhole pairs.")
        
        print("\nAll required mod files have been generated:")
        for path in [output_map_file, output_initializer_file, output_onactions_file, output_wormhole_events_file, output_mega_events_file, output_wormhole_effects_file]:
            if os.path.exists(path):
                is_dir = os.path.isdir(path)
                is_empty = not os.listdir(path) if is_dir else os.path.getsize(path) == 0
                if not (is_dir and is_empty):
                     print(f"- {os.path.relpath(path, script_dir)}")
        
        print("\nTo load your imported game, select the 'Continuum' galaxy when starting a New Game.")

    else:
        print("Could not parse critical galaxy data from the save file.")

    input("\nPress Enter to exit.")

if __name__ == "__main__":
    main()