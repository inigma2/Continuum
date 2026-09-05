"""0.7: parse Pre default empires / primitives, pick a new-polity start, emit Continuum Present snapshot."""
import random
import re
import zipfile

import continuum_cp

UNINHABITABLE_CLASSES = frozenset({
    "pc_barren", "pc_barren_cold", "pc_toxic", "pc_frozen", "pc_molten",
    "pc_gas_giant", "pc_asteroid", "pc_ice_asteroid", "pc_rare_crystal_asteroid",
    "pc_shattered", "pc_shattered_2", "pc_broken", "pc_cracked", "pc_shielded",
    "pc_ai", "pc_infested", "pc_gray_goo", "pc_egg_cracked",
    "pc_shrouded",
})
UNIQUE_BLOCK_FLAGS = frozenset({
    "guardians_artists_system", "guardians_curators_system", "guardians_traders_system",
    "salvager_enclave_system", "shroudwalker_enclave_system", "shroud_tunnel_nexus",
    "shroud_tunnel_node", "spawned_shroud_tunnel",
    "guardians_dragon_system", "guardians_technosphere_system", "guardians_wraith_system",
    "guardians_horror_system", "guardians_dreadnought_system", "guardians_hive_system",
    "guardians_fortress_system", "guardians_stellarite_system", "guardians_hatchling_system",
    "lcluster1", "lcluster", "terminal_egress", "crystal_home_system",
    "amoeba_home_system", "drone_home_system", "voidworms_system", "elderly_tiyanki_system",
    "lost_swarm_system", "wenkwort_system",
    "marauder_capital_1", "marauder_capital_2", "marauder_capital_3",
})

PRE_FTL_AGES = (
    "stone_age", "bronze_age", "iron_age", "late_medieval_age", "renaissance_age",
    "steam_age", "industrial_age", "machine_age", "atomic_age", "early_space_age",
)

FALLEN_DESIGNS = {
    "1": (  # materialist
        "NAME_Enforcer", "NAME_Savant", "NAME_Scholar", "NAME_Sage", "NAME_Cloaker",
        "NAME_Librarian", "NAME_Seeker", "NAME_FE_MATERIALIST_Citadel_1",
        "NAME_FE_MATERIALIST_Citadel_2", "NAME_FE_MATERIALIST_Citadel_3", "NAME_FE_Starbase",
    ),
    "2": (  # spiritualist
        "NAME_Cleanser", "NAME_Eternal", "NAME_Avatar", "NAME_Zealot", "NAME_Penitent",
        "NAME_Faith", "NAME_Pilgrim", "NAME_FE_SPIRITUALIST_Citadel_1",
        "NAME_FE_SPIRITUALIST_Citadel_2", "NAME_FE_SPIRITUALIST_Citadel_3", "NAME_FE_Starbase",
    ),
    "3": (  # xenophile
        "NAME_Adjuster", "NAME_Keeper", "NAME_Custodian", "NAME_Overseer", "NAME_Watcher",
        "NAME_Seeder", "NAME_Builder", "NAME_FE_XENOPHILE_Citadel_1",
        "NAME_FE_XENOPHILE_Citadel_2", "NAME_FE_XENOPHILE_Citadel_3", "NAME_FE_Starbase",
    ),
    "4": (  # xenophobe
        "NAME_Reaper", "NAME_Imperium", "NAME_Supremacy", "NAME_Glory", "NAME_Devastator",
        "NAME_Servitor", "NAME_Destiny", "NAME_FE_XENOPHOBE_Citadel_1",
        "NAME_FE_XENOPHOBE_Citadel_2", "NAME_FE_XENOPHOBE_Citadel_3", "NAME_FE_Starbase",
    ),
    "machine": (
        "NAME_Omega", "NAME_Alpha", "NAME_Beta", "NAME_Gamma", "NAME_Theta",
        "NAME_Tau", "NAME_Sigma", "NAME_FE_MACHINE_Citadel_1",
        "NAME_FE_MACHINE_Citadel_2", "NAME_FE_MACHINE_Citadel_3", "NAME_FE_Starbase",
    ),
}

FALLEN_BY_ETHIC = {
    "ethic_fanatic_materialist": "1",
    "ethic_fanatic_spiritualist": "2",
    "ethic_fanatic_xenophile": "3",
    "ethic_fanatic_xenophobe": "4",
}
# military_station_small_fallen_empire global designs (common/global_ship_designs/fallen_empire_ship_designs.txt)
FE_PLATFORM_DESIGNS = {
    "1": "NAME_Cloaker",
    "2": "NAME_Faith",
    "3": "NAME_Watcher",
    "4": "NAME_Devastator",
    "machine": "NAME_Sigma",
}
FE_SHIP_DESIGNS = {
    "1": {
        "small_ship_fallen_empire": "NAME_Sage",
        "large_ship_fallen_empire": "NAME_Scholar",
        "massive_ship_fallen_empire": "NAME_Savant",
    },
    "2": {
        "small_ship_fallen_empire": "NAME_Zealot",
        "large_ship_fallen_empire": "NAME_Avatar",
        "massive_ship_fallen_empire": "NAME_Eternal",
    },
    "3": {
        "small_ship_fallen_empire": "NAME_Overseer",
        "large_ship_fallen_empire": "NAME_Custodian",
        "massive_ship_fallen_empire": "NAME_Keeper",
    },
    "4": {
        "small_ship_fallen_empire": "NAME_Glory",
        "large_ship_fallen_empire": "NAME_Supremacy",
        "massive_ship_fallen_empire": "NAME_Imperium",
    },
    "machine": {
        "small_ship_fallen_empire": "NAME_Gamma",
        "large_ship_fallen_empire": "NAME_Beta",
        "massive_ship_fallen_empire": "NAME_Alpha",
    },
}
FE_DSC_DESIGNS = {
    "1": {
        "starbase_deep_space_citadel_1": "NAME_FE_MATERIALIST_Citadel_1",
        "starbase_deep_space_citadel_2": "NAME_FE_MATERIALIST_Citadel_2",
        "starbase_deep_space_citadel_3": "NAME_FE_MATERIALIST_Citadel_3",
    },
    "2": {
        "starbase_deep_space_citadel_1": "NAME_FE_SPIRITUALIST_Citadel_1",
        "starbase_deep_space_citadel_2": "NAME_FE_SPIRITUALIST_Citadel_2",
        "starbase_deep_space_citadel_3": "NAME_FE_SPIRITUALIST_Citadel_3",
    },
    "3": {
        "starbase_deep_space_citadel_1": "NAME_FE_XENOPHILE_Citadel_1",
        "starbase_deep_space_citadel_2": "NAME_FE_XENOPHILE_Citadel_2",
        "starbase_deep_space_citadel_3": "NAME_FE_XENOPHILE_Citadel_3",
    },
    "4": {
        "starbase_deep_space_citadel_1": "NAME_FE_XENOPHOBE_Citadel_1",
        "starbase_deep_space_citadel_2": "NAME_FE_XENOPHOBE_Citadel_2",
        "starbase_deep_space_citadel_3": "NAME_FE_XENOPHOBE_Citadel_3",
    },
    "machine": {
        "starbase_deep_space_citadel_1": "NAME_FE_MACHINE_Citadel_1",
        "starbase_deep_space_citadel_2": "NAME_FE_MACHINE_Citadel_2",
        "starbase_deep_space_citadel_3": "NAME_FE_MACHINE_Citadel_3",
    },
}
# marauder ship_size -> global design (common/global_ship_designs/marauder_ship_designs.txt)
MARAUDER_SHIP_DESIGNS = {
    "marauder_corvette": "NAME_Outrider",
    "marauder_destroyer": "NAME_Lancer",
    "marauder_cruiser": "NAME_Void_Champion",
    "marauder_galleon": "NAME_Ancestral_Glory",
}
MARAUDER_STATION_DESIGNS = {
    "marauder_station": "NAME_Warrior_Freehold",
    "marauder_void_dwelling": "NAME_Void_Dwelling",
}
MARAUDER_STATION_SIZES = frozenset(MARAUDER_STATION_DESIGNS)
MARAUDER_GLOBAL_DESIGNS = (
    "NAME_Warrior_Freehold",
    "NAME_Void_Dwelling",
    "NAME_Outrider",
    "NAME_Lancer",
    "NAME_Void_Champion",
    "NAME_Ancestral_Glory",
    "NAME_Marauder_Starbase",
)
MARAUDER_FLEET_DEFAULTS = {
    "marauder_corvette": 12,
    "marauder_destroyer": 6,
    "marauder_cruiser": 3,
    "marauder_galleon": 1,
    "marauder_station": 8,
    "marauder_void_dwelling": 6,
}
MARAUDER_SAT_DEFAULTS = {
    "marauder_corvette": 22,
    "marauder_destroyer": 14,
    "marauder_cruiser": 8,
    "marauder_galleon": 1,
    "marauder_station": 0,
    "marauder_void_dwelling": 2,
}


def is_habitable_class(planet_class):
    if not planet_class:
        return False
    pc = planet_class.strip('"')
    if pc in UNINHABITABLE_CLASSES or pc.startswith("pc_shrouded"):
        return False
    if any(s in pc for s in ("_star", "black_hole", "pulsar", "neutron", "quasar")):
        return False
    if pc in ("pc_habitat", "pc_ringworld_habitable"):
        return False
    return pc.startswith("pc_")


def _brace_section(data, key):
    needle = f"\n{key}=\n{{"
    i = data.find(needle)
    if i < 0:
        needle = f"\n{key}={{"
        i = data.find(needle)
        if i < 0:
            return ""
        start = data.find("{", i)
    else:
        start = data.find("{", i)
    depth = 0
    for j in range(start, len(data)):
        if data[j] == "{":
            depth += 1
        elif data[j] == "}":
            depth -= 1
            if depth == 0:
                return data[start + 1:j]
    return ""


def _split_numbered(section):
    parts = re.split(r"\n\t(\d+)=\n", section)
    out = {}
    for k in range(1, len(parts) - 1, 2):
        out[parts[k]] = parts[k + 1]
    return out


def _split_top_entries(section):
    """Brace-match `\\tID=\\n\\t{...}` so nested `\\tN=` inside a country is not a new country."""
    out = {}
    hdr = re.compile(r"\t(\d+)=\s*\n\t\{")
    i = 0
    n = len(section)
    while True:
        m = hdr.search(section, i)
        if not m:
            break
        cid = m.group(1)
        brace_start = section.find("{", m.start())
        depth = 0
        j = brace_start
        while j < n:
            ch = section[j]
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    out[cid] = section[brace_start + 1:j]
                    i = j + 1
                    break
            j += 1
        else:
            break
    return out


def set_loc_data(loc_data):
    continuum_cp.set_loc_data(loc_data)


def _name_from_block(body):
    i = body.find("\n\t\tname=")
    if i < 0:
        i = body.find("\n\tname=")
    if i < 0:
        m = re.search(r'key="([^"]+)"', body[:1200])
        return continuum_cp.name_from_keys([m.group(1)]) if m else "Unknown"
    start = body.find("{", i)
    if start < 0:
        m = re.search(r'key="([^"]+)"', body[i:i + 400])
        return continuum_cp.name_from_keys([m.group(1)]) if m else "Unknown"
    depth = 0
    blob = ""
    for j in range(start, len(body)):
        if body[j] == "{":
            depth += 1
        elif body[j] == "}":
            depth -= 1
            if depth == 0:
                blob = body[start + 1:j]
                break
    keys = re.findall(r'key="([^"]+)"', blob)
    if not keys:
        return "Unknown"
    return continuum_cp.name_from_keys(keys)


def display_empire_name(name):
    """Readable country name. Loc-resolved; never leaves EMPIRE_DESIGN / SPEC_ / AofB debris."""
    s = (name or "Empire").strip()
    if not s:
        return "Empire"
    if " " in s or s == continuum_cp.display_key(s):
        if s.upper().startswith("EMPIRE DESIGN"):
            rest = s[13:].strip(" :_-")
            return rest[:1].upper() + rest[1:] if rest else "Empire"
        return s
    return continuum_cp.display_key(s) or s


def script_name(name):
    if not name:
        return '"Unknown"'
    s = str(name)
    if re.match(r"^(NAME_|SPEC_|PRESCRIPTED_|EMPIRE_DESIGN_)[A-Za-z0-9_]+$", s):
        return s
    return f'"{s.replace(chr(34), "")}"'


def parse_species_db(data):
    section = _brace_section(data, "species_db")
    species = {}
    for sid, body in _split_numbered(section).items():
        if "class=" not in body:
            continue
        cls = re.search(r'class="([^"]+)"', body)
        portrait = re.search(r'portrait="([^"]+)"', body)
        namelist = re.search(r'name_list="([^"]+)"', body)
        name = _name_from_block(body)
        traits = re.findall(r'trait="([^"]+)"', body)
        species[sid] = {
            "id": sid,
            "name": name,
            "class": cls.group(1) if cls else "MAM",
            "portrait": portrait.group(1) if portrait else "human",
            "namelist": namelist.group(1) if namelist else "MAM1",
            "traits": traits,
        }
    return species


def parse_colony_planet_map(data):
    """4.x owned_planets/capital are colony IDs. carrier.reference is the planet ID."""
    out = {}
    for cid, body in _split_top_entries(_brace_section(data, "colony")).items():
        m = re.search(r"carrier=\s*\{\s*type=planet\s*reference=(\d+)", body)
        if m:
            out[str(cid)] = m.group(1)
    return out


def _system_of_colony(token, colony_to_planet, p2s):
    pid = (colony_to_planet or {}).get(str(token), str(token))
    return p2s.get(str(pid))


def parse_countries(data):
    section = _brace_section(data, "country")
    countries = []
    keep = ("default", "primitive", "fallen_empire", "dormant_marauders", "enclave")
    for cid, body in _split_top_entries(section).items():
        typ = re.search(r'\n\t\ttype="([^"]+)"', body)
        if not typ:
            continue
        t = typ.group(1)
        if t not in keep:
            continue
        ss = re.search(r"starting_system=(\d+)", body)
        owned = re.search(r"owned_planets=\s*\n\t\t\{\s*\n\t\t\t([^}]+)\}", body)
        cap = re.search(r"\n\t\tcapital=(\d+)", body)
        founder = re.search(r"founder_species_ref=(\d+)", body)
        ethics = re.findall(r'ethic="([^"]+)"', body)
        gov_i = body.find("government=")
        gov_chunk = body[gov_i:gov_i + 2500] if gov_i >= 0 else body[:2500]
        civics = re.findall(r'"(civic_[^"]+)"', gov_chunk)
        if not civics:
            civics_block = re.search(r"civics=\s*\{([^}]+)\}", body)
            civics = re.findall(r'"([^"]+)"', civics_block.group(1)) if civics_block else []
        auth = re.search(r'authority="([^"]+)"', body)
        gfx = re.search(r'graphical_culture="([^"]+)"', body)
        origin = re.search(r'origin="([^"]+)"', body)
        icon_cat = re.search(r'icon=\s*\{\s*category="([^"]+)"\s*file="([^"]+)"', body)
        bg = re.search(r'background=\s*\{\s*category="([^"]+)"\s*file="([^"]+)"', body)
        colors = re.search(r'colors=\s*\{\s*"([^"]+)"\s*"([^"]+)"\s*"([^"]+)"\s*"([^"]+)"', body)
        fe_n = re.search(r"fallen_empire_(\d+)=", body)
        mar_n = re.search(r"\bmarauder_(\d+)=", body)
        age = None
        for a in PRE_FTL_AGES:
            if re.search(rf"\b{a}=", body):
                age = a
                break
        countries.append({
            "id": cid,
            "type": t,
            "name": _name_from_block(body),
            "starting_system": ss.group(1) if ss else None,
            "capital": cap.group(1) if cap else None,
            "owned_planets": owned.group(1).split() if owned else [],
            "founder_species": founder.group(1) if founder else None,
            "ethics": ethics,
            "civics": [c for c in civics if c.startswith("civic_")],
            "authority": auth.group(1) if auth else "auth_oligarchic",
            "graphical_culture": gfx.group(1) if gfx else "mammalian_01",
            "origin": origin.group(1) if origin else None,
            "fallen_n": fe_n.group(1) if fe_n else None,
            "marauder_n": mar_n.group(1) if mar_n else None,
            "pre_ftl_age": age,
            "flag_icon_cat": icon_cat.group(1) if icon_cat else "special",
            "flag_icon": icon_cat.group(2) if icon_cat else "pirate_flag.dds",
            "flag_bg_cat": bg.group(1) if bg else "backgrounds",
            "flag_bg": bg.group(2) if bg else "00_solid.dds",
            "flag_colors": [c.strip().strip('"') for c in colors.groups()] if colors else ["red", "black", "black", "null"],
        })
    return countries


STARBASE_LEVELS = {
    "starbase_level_outpost": "starbase_outpost",
    "starbase_level_starport": "starbase_starport",
    "starbase_level_starhold": "starbase_starhold",
    "starbase_level_starfortress": "starbase_starfortress",
    "starbase_level_citadel": "starbase_citadel",
    "starbase_level_deep_space_citadel_1": "starbase_deep_space_citadel_1",
    "starbase_level_deep_space_citadel_2": "starbase_deep_space_citadel_2",
    "starbase_level_deep_space_citadel_3": "starbase_deep_space_citadel_3",
    "starbase_level_marauder": "starbase_marauder",
}
STARBASE_SHORT = {
    "starbase_outpost": "outpost",
    "starbase_starport": "starport",
    "starbase_starhold": "starhold",
    "starbase_starfortress": "starfortress",
    "starbase_citadel": "citadel",
    "starbase_deep_space_citadel_1": "dsc1",
    "starbase_deep_space_citadel_2": "dsc2",
    "starbase_deep_space_citadel_3": "dsc",
    "starbase_marauder": "marauder",
}
STARBASE_TECHS = {
    "starbase_outpost": ["tech_starbase_1"],
    "starbase_starport": ["tech_starbase_1", "tech_starbase_2"],
    "starbase_starhold": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3"],
    "starbase_starfortress": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3", "tech_starbase_4"],
    "starbase_citadel": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3", "tech_starbase_4", "tech_starbase_5"],
    "starbase_deep_space_citadel_1": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3", "tech_deep_space_citadel"],
    "starbase_deep_space_citadel_2": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3", "tech_deep_space_citadel"],
    "starbase_deep_space_citadel_3": ["tech_starbase_1", "tech_starbase_2", "tech_starbase_3", "tech_starbase_4", "tech_starbase_5", "tech_deep_space_citadel"],
}
# Prefer DSC over a co-located outpost (Pre DSC systems also have an outpost).
STARBASE_RANK = {
    "starbase_outpost": 1,
    "starbase_starport": 2,
    "starbase_starhold": 3,
    "starbase_starfortress": 4,
    "starbase_citadel": 5,
    "starbase_deep_space_citadel_1": 6,
    "starbase_deep_space_citadel_2": 7,
    "starbase_deep_space_citadel_3": 8,
    "starbase_marauder": 5,
}
FLEET_DESIGNS = (
    "corvette", "frigate", "destroyer", "cruiser", "battleship", "titan",
    "small_ship_fallen_empire", "large_ship_fallen_empire", "massive_ship_fallen_empire",
    "military_station_small_fallen_empire",
    "marauder_corvette", "marauder_destroyer", "marauder_cruiser", "marauder_galleon",
    "marauder_station", "marauder_void_dwelling",
)
FLEET_MAX_PER_SIZE = 40
# Spawned via _fe_defense_platforms / _marauder_presence, not the mobile home fleet.
_FLEET_BLOCK_SKIP = frozenset({
    "military_station_small_fallen_empire",
    "small_ship_fallen_empire", "large_ship_fallen_empire", "massive_ship_fallen_empire",
    "marauder_corvette", "marauder_destroyer", "marauder_cruiser", "marauder_galleon",
    "marauder_station", "marauder_void_dwelling",
})


def _id_list(body, key):
    m = re.search(rf"\n\t\t{key}=\s*\n\t\t\{{\s*\n\t\t\t([^}}]+)\}}", body)
    if not m:
        return []
    return [t for t in m.group(1).split() if t.isdigit()]


def _owned_fleet_ids(body):
    ids = []
    start = 0
    while True:
        i = body.find("owned_fleets=", start)
        if i < 0:
            break
        brace = body.find("{", i)
        if brace < 0:
            break
        depth = 0
        for k in range(brace, len(body)):
            if body[k] == "{":
                depth += 1
            elif body[k] == "}":
                depth -= 1
                if depth == 0:
                    ids.extend(re.findall(r"fleet=(\d+)", body[brace:k + 1]))
                    start = k + 1
                    break
        else:
            break
    return ids


def _mgr_starbases(data):
    mgr = _brace_section(data, "starbase_mgr")
    hdr = re.compile(r"\n\t\t(\d+)=\n\t\t\{")
    out = {}
    pos = 0
    n = len(mgr)
    while True:
        m = hdr.search(mgr, pos)
        if not m:
            break
        bid = m.group(1)
        start = mgr.find("{", m.start())
        depth = 0
        k = start
        while k < n:
            if mgr[k] == "{":
                depth += 1
            elif mgr[k] == "}":
                depth -= 1
                if depth == 0:
                    blk = mgr[start:k + 1]
                    lv = re.search(r'level="([^"]+)"', blk)
                    st = re.search(r"\n\t\t\tstation=(\d+)", blk)
                    size = STARBASE_LEVELS.get(lv.group(1) if lv else "")
                    out[bid] = {
                        "size": size,
                        "station": st.group(1) if st else None,
                        "level": lv.group(1) if lv else "",
                    }
                    pos = k + 1
                    break
            k += 1
        else:
            break
    return out


def attach_snapshot(data, countries, colony_to_planet, p2s):
    """Starbase territory/size, techs, pops, military ship counts from Pre."""
    by_id = {c["id"]: c for c in countries}
    fleet_owner = {}
    for c in countries:
        for fid in _id_list(c.get("_body") or "", "fleets"):
            fleet_owner[fid] = c["id"]
    # fleets live on original country bodies — stash during parse
    # fallback: re-read country section
    csec = _split_top_entries(_brace_section(data, "country"))
    fleet_owner = {}
    for cid, body in csec.items():
        if cid not in by_id:
            continue
        for fid in _owned_fleet_ids(body):
            if fid not in fleet_owner:
                fleet_owner[fid] = cid
        for fid in _id_list(body, "fleets"):
            if fid not in fleet_owner:
                fleet_owner[fid] = cid
        ts = re.search(r"tech_status=\s*\{", body)
        techs = []
        if ts:
            start = body.find("{", ts.start())
            depth = 0
            for k in range(start, len(body)):
                if body[k] == "{":
                    depth += 1
                elif body[k] == "}":
                    depth -= 1
                    if depth == 0:
                        techs = re.findall(r'technology="(tech_[^"]+)"', body[start:k + 1])
                        break
        by_id[cid]["techs"] = techs
        by_id[cid]["starbases"] = {}
        by_id[cid]["dsc"] = {}
        by_id[cid]["fleet_mix"] = {}
        by_id[cid]["system_fleets"] = {}
        by_id[cid]["colony_pop"] = {}

    g = _split_top_entries(_brace_section(data, "galactic_object"))
    sb_sys = {}
    for sid, body in g.items():
        m = re.search(r"starbases=\s*\{([^}]*)\}", body)
        if not m:
            continue
        for t in m.group(1).split():
            if t.isdigit() and t != "4294967295":
                sb_sys[t] = str(sid)

    designs = {}
    for did, body in _split_top_entries(_brace_section(data, "ship_design")).items():
        m = re.search(r'ship_size="([^"]+)"', body)
        if m:
            designs[did] = m.group(1)
    ship_fleet = {}
    for _sid, body in _split_top_entries(_brace_section(data, "ships")).items():
        fl = re.search(r"\n\t\tfleet=(\d+)", body)
        ds = re.search(r"\n\t\t\tdesign=(\d+)", body)
        if not ds:
            ds = re.search(r"\ndesign=(\d+)", body)
        coord = re.search(r"coordinate=\s*\{([^}]+)\}", body)
        ss = re.search(r"origin=(\d+)", coord.group(1)) if coord else None
        if fl:
            ship_fleet[_sid] = fl.group(1)
        if not fl or not ds:
            continue
        owner = fleet_owner.get(fl.group(1))
        sz = designs.get(ds.group(1))
        if owner in by_id and sz in FLEET_DESIGNS:
            mix = by_id[owner]["fleet_mix"]
            mix[sz] = mix.get(sz, 0) + 1
            if ss:
                slot = by_id[owner]["system_fleets"].setdefault(ss.group(1), {})
                slot[sz] = slot.get(sz, 0) + 1

    for sbid, info in _mgr_starbases(data).items():
        size = info.get("size")
        if not size:
            continue
        station = info.get("station")
        # 4.4 starbase_mgr station= is the starbase SHIP id, not the fleet id.
        fleet_id = ship_fleet.get(station) or station
        owner = fleet_owner.get(fleet_id)
        sys_id = sb_sys.get(sbid)
        if owner and sys_id and owner in by_id:
            if str(size).startswith("starbase_deep_space_citadel"):
                by_id[owner].setdefault("dsc", {})[sys_id] = size
                continue
            prev = by_id[owner]["starbases"].get(sys_id)
            if prev is None or STARBASE_RANK.get(size, 0) >= STARBASE_RANK.get(prev, 0):
                by_id[owner]["starbases"][sys_id] = size

    cols = _split_top_entries(_brace_section(data, "colony"))
    for c in countries:
        pops = {}
        for col in c.get("owned_planets") or []:
            pid = colony_to_planet.get(str(col))
            blk = cols.get(str(col), "")
            n = re.search(r"num_sapient_pops=(\d+)", blk)
            if pid:
                pops[str(pid)] = max(100, min(int(n.group(1)) if n else 1000, 12000))
        c["colony_pop"] = pops
        for sz, n in list((c.get("fleet_mix") or {}).items()):
            if sz in MARAUDER_STATION_SIZES or sz in (
                "military_station_small_fallen_empire",
                "small_ship_fallen_empire",
                "large_ship_fallen_empire",
                "massive_ship_fallen_empire",
            ):
                continue
            if n > FLEET_MAX_PER_SIZE:
                c["fleet_mix"][sz] = FLEET_MAX_PER_SIZE
    continuum_cp.attach_cp_layers(data, countries)
    return countries


def names_from_galaxy(galaxy_data):
    """Planet id → display name used in initializers (create_colony overwrites this)."""
    out = {}
    for sys in galaxy_data:
        root = sys.get("hierarchy_root")
        queue = [root] if root else []
        while queue:
            body = queue.pop(0)
            queue.extend(body.get("children") or [])
            bid = body.get("id")
            nm = body.get("name")
            if bid is not None and nm:
                out[str(bid)] = nm
    return out


def attach_colony_names(countries, name_by_pid, colony_to_planet=None):
    colony_to_planet = colony_to_planet or {}
    for c in countries:
        names = {}
        for pid in c.get("colony_pop") or {}:
            nm = name_by_pid.get(str(pid))
            if nm:
                names[str(pid)] = nm
        for token in c.get("owned_planets") or []:
            pid = colony_to_planet.get(str(token))
            nm = name_by_pid.get(str(pid)) if pid else None
            if nm:
                names[str(pid)] = nm
        c["colony_names"] = names
        if c.get("type") != "default":
            continue
        cap_col = c.get("capital")
        cap_pid = colony_to_planet.get(str(cap_col)) if cap_col is not None else None
        if not cap_pid:
            pops = list((c.get("colony_pop") or {}).keys())
            cap_pid = str(pops[0]) if pops else None
        if not cap_pid or str(cap_pid) not in names:
            continue
        nm = names[str(cap_pid)]
        if nm.startswith("NAME_") or nm.startswith("SPEC_") or nm.startswith("%"):
            pretty = nm.replace("SPEC_", "").replace("NAME_", "").replace("%", "").replace("_", " ")
        else:
            pretty = nm
        if pretty and not pretty.endswith(" Prime"):
            names[str(cap_pid)] = pretty + " Prime"
    return countries


def _plain_name(nm):
    if not nm:
        return ""
    s = str(nm)
    if s.startswith("NAME_") or s.startswith("SPEC_"):
        s = s.replace("NAME_", "").replace("SPEC_", "").replace("_", " ")
    if s.endswith(" Prime"):
        s = s[: -len(" Prime")]
    return s.strip()


def _with_prime(nm):
    if not nm:
        return nm
    if str(nm).endswith(" Prime"):
        return nm
    pretty = _plain_name(nm)
    return (pretty + " Prime") if pretty else nm


def _similar_name(nm, base):
    a, b = _plain_name(nm).lower(), _plain_name(base).lower()
    if not a or not b:
        return False
    if a == b:
        return True
    if a.startswith(b + " ") or a.startswith(b + "-"):
        return True
    if b.startswith(a + " ") or b.startswith(a + "-"):
        return True
    return False


def apply_home_prime(galaxy_data, defaults, capitals):
    """Prime homeworld planets only. System names stay as in Pre (Woh, not Woh Prime)."""
    by_id = {str(s.get("id")): s for s in galaxy_data}
    for c in defaults:
        cap = capitals.get(c["id"])
        if not cap:
            continue
        sys = by_id.get(str(cap))
        if not sys:
            continue
        names = c.get("colony_names") or {}
        hw = ""
        for nm in names.values():
            if str(nm).endswith(" Prime"):
                hw = _plain_name(nm)
                break
        if not hw and names:
            hw = _plain_name(next(iter(names.values())))
        if not hw:
            continue
        root = sys.get("hierarchy_root")
        queue = [root] if root else []
        while queue:
            body = queue.pop(0)
            queue.extend(body.get("children") or [])
            nm = body.get("name")
            # Stars keep Cyban A / Woh. Only the homeworld body gets Prime.
            if body.get("body_type") == "star":
                continue
            if nm and _similar_name(nm, hw):
                body["name"] = _with_prime(nm)


def planet_to_system_map(stars):
    mapping = {}
    for sid, star in stars.items():
        for pid in star.get("planet_ids") or []:
            mapping[str(pid)] = str(sid)
    return mapping


def system_habitables(galaxy_data, planets):
    out = {}
    for sys in galaxy_data:
        sid = str(sys.get("id"))
        hab = []
        root = sys.get("hierarchy_root")
        queue = [root] if root else []
        while queue:
            body = queue.pop(0)
            queue.extend(body.get("children") or [])
            pc = body.get("planet_class") or ""
            if is_habitable_class(pc):
                hab.append(body)
        if not hab:
            for pid in sys.get("planet_ids") or []:
                pc = (planets.get(pid) or planets.get(str(pid)) or {}).get("planet_class", "")
                if is_habitable_class(pc):
                    hab.append({"id": pid, "planet_class": pc, "name": (planets.get(pid) or {}).get("name")})
        out[sid] = hab
    return out


def blocked_systems(galaxy_data):
    blocked = set()
    for sys in galaxy_data:
        flags = set(sys.get("flags") or [])
        if flags & UNIQUE_BLOCK_FLAGS:
            blocked.add(str(sys.get("id")))
        if any(str(fl).startswith("guardians_") or str(fl).startswith("lcluster") for fl in flags):
            blocked.add(str(sys.get("id")))
    return blocked


def ownership_from_planets(countries, p2s, colony_to_planet=None, types=("default",)):
    owned = {}
    capitals = {}
    for c in countries:
        if c["type"] not in types:
            continue
        systems = []

        def add_colony(token):
            sid = _system_of_colony(token, colony_to_planet, p2s)
            if sid and sid not in systems:
                systems.append(sid)
            return sid

        for token in c.get("owned_planets") or []:
            add_colony(token)
        cap_sid = add_colony(c["capital"]) if c.get("capital") is not None else None
        if c.get("starting_system") and str(c["starting_system"]) not in systems:
            systems.insert(0, str(c["starting_system"]))
        owned[c["id"]] = systems
        capitals[c["id"]] = (
            str(cap_sid) if cap_sid else (
                str(c["starting_system"]) if c.get("starting_system") else (systems[0] if systems else None)
            )
        )
    return owned, capitals


def nudge_borders(owned, capitals, hyperlanes, blocked, reserved, rng):
    """0–2 jump frontier shift. Capitals and reserved systems stay put. No extinctions."""
    owned = {cid: list(syss) for cid, syss in owned.items()}
    neighbors = hyperlanes
    strength = {cid: max(1, len(syss)) for cid, syss in owned.items()}
    claimed = {s for syss in owned.values() for s in syss}

    def owner_of(sid):
        for cid, syss in owned.items():
            if sid in syss:
                return cid
        return None

    for _round in range(2):
        order = sorted(owned.keys(), key=lambda c: -strength[c])
        for cid in order:
            frontier = []
            for sid in owned[cid]:
                for nb in neighbors.get(sid, []):
                    if nb in blocked or nb in reserved:
                        continue
                    frontier.append(nb)
            unowned = [s for s in frontier if owner_of(s) is None and s not in claimed]
            rng.shuffle(unowned)
            take = 1 if strength[cid] >= 3 else (1 if rng.random() < 0.55 else 0)
            for sid in unowned[:take]:
                owned[cid].append(sid)
                claimed.add(sid)
            # peel a weak edge
            for sid in frontier:
                other = owner_of(sid)
                if not other or other == cid:
                    continue
                if sid == capitals.get(other) or sid in reserved:
                    continue
                if strength[cid] > strength[other] * 1.4 and rng.random() < 0.18:
                    if len(owned[other]) <= 1:
                        continue
                    owned[other].remove(sid)
                    owned[cid].append(sid)
                    break
        strength = {cid: max(1, len(syss)) for cid, syss in owned.items()}
    return owned


CRISIS_COUNTRY_TYPES = (
    "swarm", "extradimensional", "ai_empire", "gray_goo", "synth_queen", "formless",
)

# 4.4: these keys exist but fail give_technology on gestalt/machine countries.
GESTALT_SKIP_TECHS = frozenset({
    "tech_interplanetary_commerce",
    "tech_holo_entertainment",
    "tech_critter_feeder",
})
ALWAYS_SKIP_TECHS = frozenset({
    "tech_critter_feeder",
})


def detect_crisis(data):
    section = _brace_section(data, "country")
    for t in CRISIS_COUNTRY_TYPES:
        if f'type="{t}"' in section:
            return True
    return False


def species_class_of(emp, species):
    sp = species.get(str(emp.get("founder_species"))) or {}
    return sp.get("class") or "MAM"


def is_machine(emp, sp):
    cls = (sp or {}).get("class", "")
    auth = emp.get("authority") or ""
    return cls in ("MACHINE", "ROBOT") or auth == "auth_machine_intelligence"


def is_hive(emp, sp):
    if is_machine(emp, sp):
        return False
    auth = emp.get("authority") or ""
    ethics = emp.get("ethics") or []
    return auth == "auth_hive_mind" or "ethic_gestalt_consciousness" in ethics


def classify_systems(owned, capitals, lanes, blocked, habitables, countries, species):
    owner_of = {}
    for cid, syss in owned.items():
        for sid in syss:
            owner_of[sid] = cid
    class_of = {}
    for c in countries:
        if c["type"] == "default":
            class_of[c["id"]] = species_class_of(c, species)

    tags = {}
    spawn = {}

    def add_tag(sid, flag):
        lst = tags.setdefault(sid, [])
        if flag not in lst:
            lst.append(flag)

    for sid, hab in habitables.items():
        if not hab or sid in blocked:
            continue
        cid = owner_of.get(sid)
        nbs = lanes.get(sid) or []
        if cid is None:
            add_tag(sid, "continuum_unowned")
            near = []
            for nb in nbs:
                oc = owner_of.get(nb)
                if oc in class_of:
                    near.append(class_of[oc])
            near = sorted(set(near))
            mods = [{"add": 40, "class": cl} for cl in near]
            spawn[sid] = {"base": 8, "modifiers": mods}
            for cl in near:
                add_tag(sid, f"continuum_near_{cl}")
            continue
        cl = class_of.get(cid, "MAM")
        add_tag(sid, f"continuum_species_{cl}")
        is_cap = sid == capitals.get(cid)
        is_border = any(owner_of.get(nb) != cid for nb in nbs)
        if is_cap:
            add_tag(sid, "continuum_core")
            continue
        if is_border:
            add_tag(sid, "continuum_border")
            other_emp = any(
                owner_of.get(nb) in class_of and owner_of.get(nb) != cid for nb in nbs
            )
            if other_emp:
                add_tag(sid, "continuum_front")
                spawn[sid] = {"base": 6, "modifiers": [{"add": 100, "class": cl}]}
            else:
                spawn[sid] = {"base": 2, "modifiers": [{"add": 35, "class": cl}]}
        else:
            add_tag(sid, "continuum_core")

    prim_classes = {}
    for c in countries:
        if c["type"] != "primitive":
            continue
        sid = c.get("system_id")
        if not sid or sid in blocked:
            continue
        prim_classes.setdefault(sid, set()).add(species_class_of(c, species))
    for sid, classes in prim_classes.items():
        add_tag(sid, "continuum_was_primitive")
        prev = spawn.get(sid, {"base": 0, "modifiers": []})
        mods = list(prev.get("modifiers") or [])
        have = {(int(m.get("add", 0)), m.get("class")) for m in mods}
        for cl in sorted(classes):
            add_tag(sid, f"continuum_species_{cl}")
            key = (50, cl)
            if key not in have:
                mods.append({"add": 50, "class": cl})
                have.add(key)
        spawn[sid] = {"base": max(int(prev.get("base") or 0), 6), "modifiers": mods}

    return tags, spawn


def format_spawn_weight(spec):
    if not spec:
        return ""
    parts = [f"base = {int(spec.get('base', 1))}"]
    seen = set()
    for m in spec.get("modifiers") or []:
        add = int(m.get("add", 0))
        cl = m.get("class")
        key = (add, cl)
        if not cl or key in seen:
            continue
        seen.add(key)
        # static_galaxy_scenario evaluates this with the spawning country as THIS, not FROM.
        parts.append(f"modifier = {{ add = {add} is_species_class = {cl} }}")
    return " spawn_weight = { " + " ".join(parts) + " }"


def _prompt(options, header):
    print(f"\n{header}")
    for i, (_k, label) in enumerate(options, 1):
        print(f"  [{i}] {label}")
    while True:
        raw = input("Enter a selection: ").strip().lower()
        if raw == "q":
            return None
        try:
            n = int(raw)
            if 1 <= n <= len(options):
                return options[n - 1][0]
        except ValueError:
            pass
        print("Invalid number.")


def choose_start(countries, species, owned, capitals, habitables, blocked, hyperlanes, argv_start=None):
    defaults = [c for c in countries if c["type"] == "default" and owned.get(c["id"])]
    primitives = [c for c in countries if c["type"] == "primitive"]
    species_by_id = species

    def species_label(sid):
        sp = species_by_id.get(sid) or species_by_id.get(str(sid))
        return sp["name"] if sp else f"species {sid}"

    argv_map = {
        "new": "new",
        "primitive": "primitive",
        "same": "same",
        "civil": "same",
        "civil_war": "same",
        "colony": "same",
        "random": "same",
    }
    kind = argv_map.get(argv_start) if argv_start else None
    if kind is None:
        kind = _prompt(
            [
                ("same", "New empire, same species as a Pre empire"),
                ("new", "New empire, new species (unowned habitable)"),
                ("primitive", "New empire, formerly a Pre pre-FTL species"),
            ],
            "0.7 — what kind of start? You are always a new political entity.",
        )
        if kind is None:
            return None

    plan = {
        "kind": kind,
        "player_system": None,
        "remnant_cid": None,
        "devastation_system": None,
        "intro_key": "continuum_intro_new",
        "intro_name": "",
        "reserved": set(),
    }

    if kind == "new":
        candidates = [
            sid for sid, hab in habitables.items()
            if hab and sid not in blocked and not any(sid in syss for syss in owned.values())
        ]
        if not candidates:
            candidates = [sid for sid, hab in habitables.items() if hab and sid not in blocked]
        plan["player_system"] = candidates[0] if candidates else None
        plan["intro_key"] = "continuum_intro_new"
        return plan

    if kind == "primitive":
        opts = []
        for c in primitives:
            pids = c.get("owned_planets") or []
            sid = None
            # filled later if we pass p2s via owned planets on country - primitives used planet ids
            opts.append((c["id"], f"{c['name']} ({species_label(c.get('founder_species'))})"))
        if not opts:
            print("No primitives in this save. Using a new-species start.")
            return choose_start(countries, species, owned, capitals, habitables, blocked, hyperlanes, argv_start="new")
        pick = opts[0][0] if argv_start == "primitive" else _prompt(opts, "Which Pre primitive species?")
        if pick is None:
            return None
        prim = next(c for c in primitives if c["id"] == pick)
        # starting system from first owned planet is resolved by caller via p2s stored on country
        plan["player_system"] = prim.get("system_id")
        plan["intro_key"] = "continuum_intro_primitive"
        plan["intro_name"] = prim["name"]
        plan["steal_from_overlord"] = True
        return plan

    # same species
    by_species = {}
    for c in defaults:
        sid = str(c.get("founder_species"))
        by_species.setdefault(sid, []).append(c)
    spec_opts = [(sid, f"{species_label(sid)} — {', '.join(x['name'] for x in emps)}") for sid, emps in by_species.items()]
    spec_opts.sort(key=lambda x: x[1].lower())
    if argv_start:
        richest = max(defaults, key=lambda c: len(owned.get(c["id"], [])))
        spec_id = str(richest.get("founder_species"))
    else:
        spec_id = _prompt(spec_opts, "Which Pre species are you?")
    if spec_id is None:
        return None
    empires = by_species[spec_id]
    # prefer player-like: most systems
    empires = sorted(empires, key=lambda c: -len(owned.get(c["id"], [])))
    emp = empires[0]
    loc_opts = []
    capital = capitals.get(emp["id"])
    other_hab = [s for s in owned.get(emp["id"], []) if s != capital and habitables.get(s) and s not in blocked]
    if capital and habitables.get(capital) and other_hab:
        loc_opts.append(("civil_war", f"Civil war / ceasefire — you hold {capital}'s capital, remnant on a colony"))
    if other_hab:
        loc_opts.append(("colony", "Former colony — old empire keeps the capital, you are independent"))
    loc_opts.append(("random", "Random unowned world (refugee / forgotten)"))
    loc = {
        "civil": "civil_war",
        "civil_war": "civil_war",
        "same": "civil_war",
        "colony": "colony",
        "random": "random",
    }.get(argv_start)
    if loc is None:
        loc = _prompt(loc_opts, f"Start location as a new {species_label(spec_id)} polity:")
    if loc is None:
        return None
    if loc == "civil_war" and not any(k == "civil_war" for k, _ in loc_opts):
        loc = loc_opts[0][0]
    plan["intro_name"] = emp["name"]
    plan["source_cid"] = emp["id"]
    plan["kind"] = loc
    if loc == "civil_war":
        plan["player_system"] = capital
        plan["remnant_cid"] = emp["id"]
        plan["devastation_system"] = capital
        plan["intro_key"] = "continuum_intro_civil_war"
        plan["reserved"].add(capital)
        if other_hab:
            plan["reserved"].add(other_hab[0])
            plan["remnant_capital"] = other_hab[0]
    elif loc == "colony":
        plan["player_system"] = other_hab[0]
        plan["intro_key"] = "continuum_intro_colony"
        plan["reserved"].add(other_hab[0])
    else:
        free = [sid for sid, hab in habitables.items() if hab and sid not in blocked and not any(sid in syss for syss in owned.values())]
        plan["player_system"] = free[0] if free else capital
        plan["intro_key"] = "continuum_intro_random"
    return plan


def apply_player_hole(owned, capitals, plan):
    """Remove player start (and primitive steal) from NPC ownership."""
    ps = plan.get("player_system")
    remnant_keep = plan.get("remnant_capital")
    source = plan.get("source_cid") or plan.get("remnant_cid")
    for cid, syss in list(owned.items()):
        if ps in syss:
            if plan.get("kind") == "civil_war" and cid == source:
                owned[cid] = [s for s in syss if s != ps]
                if remnant_keep and remnant_keep not in owned[cid] and remnant_keep in syss:
                    pass
            else:
                owned[cid] = [s for s in syss if s != ps]
        if not owned[cid] and cid != source:
            # keep at least capital if we emptied someone via primitive steal of their only world
            cap = capitals.get(cid)
            if cap and cap != ps:
                owned[cid] = [cap]
    if plan.get("kind") == "civil_war" and source:
        keep = [s for s in owned.get(source, []) if s != ps]
        if remnant_keep and remnant_keep not in keep:
            keep.append(remnant_keep)
        owned[source] = keep
        if remnant_keep:
            capitals[source] = remnant_keep
    return owned, capitals


def load_empire_save_data(save_path):
    with zipfile.ZipFile(save_path, "r") as z:
        data = z.read("gamestate").decode("utf-8", "replace")
    return parse_countries(data), parse_species_db(data), detect_crisis(data), data, parse_colony_planet_map(data)


def hyperlane_map(galaxy_data):
    lanes = {}
    for sys in galaxy_data:
        sid = str(sys.get("id"))
        lanes[sid] = [str(t) for t in (sys.get("hyperlanes") or [])]
    return lanes


def attach_primitive_systems(countries, p2s, colony_to_planet=None):
    for c in countries:
        if c["type"] != "primitive":
            continue
        for token in c.get("owned_planets") or []:
            sid = _system_of_colony(token, colony_to_planet, p2s)
            if sid:
                c["system_id"] = sid
                break


def script_token_ok(value, allowed):
    return (not allowed) or value in allowed


def write_opinion_file(path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("""opinion_continuum_ceasefire = {
	opinion = {
		base = -80
	}
	decay = {
		base = 0.25
	}
}
""")


def _tech_effect_lines(emp, tech_keys, extra=None):
    techs = list(emp.get("techs") or [])
    seen = set(techs)
    for t in extra or []:
        if t not in seen:
            techs.append(t)
            seen.add(t)
    for size in list((emp.get("starbases") or {}).values()) + list((emp.get("dsc") or {}).values()):
        for t in STARBASE_TECHS.get(size, []):
            if t not in seen:
                techs.append(t)
                seen.add(t)
    if "tech_corvettes" not in seen:
        techs.append("tech_corvettes")
    gestalt = (
        emp.get("authority") in ("auth_hive_mind", "auth_machine_intelligence")
        or "ethic_gestalt_consciousness" in (emp.get("ethics") or [])
        or any("machine" in c or "hive" in c for c in (emp.get("civics") or []))
    )
    skip = set(ALWAYS_SKIP_TECHS)
    if gestalt:
        skip |= GESTALT_SKIP_TECHS
    lines = []
    for t in techs:
        if tech_keys and t not in tech_keys:
            continue
        if t in skip:
            continue
        lines.append(f"					give_technology = {{ tech = {t} message = no }}")
    lines.append("					refresh_auto_generated_ship_designs = yes")
    return "\n".join(lines) if lines else "					refresh_auto_generated_ship_designs = yes"


def _starbase_blocks(prefix, idx, owner, emp=None):
    dsc = FE_DSC_DESIGNS.get(_fe_key(emp) if emp else "1") or FE_DSC_DESIGNS["1"]
    parts = []
    for size, short in STARBASE_SHORT.items():
        if size == "starbase_marauder":
            continue
        design = dsc.get(size)
        if str(size).startswith("starbase_deep_space_citadel"):
            if not emp or emp.get("type") not in ("fallen_empire",):
                continue
            citadel_design = design or "NAME_FE_XENOPHOBE_Citadel_3"
            parts.append(f"""			every_system = {{
				limit = {{
					has_star_flag = {prefix}_{idx}_{short}
					exists = event_target:{owner}
				}}
				event_target:{owner} = {{
					give_technology = {{ tech = tech_deep_space_citadel message = no }}
					create_ship_design = {{ design = "{citadel_design}" }}
					add_ship_design = last_created_design
					add_global_ship_design = "{citadel_design}"
				}}
				if = {{
					limit = {{ NOT = {{ exists = starbase }} }}
					create_starbase = {{
						size = starbase_outpost
						owner = event_target:{owner}
					}}
				}}
				random_system_planet = {{
					limit = {{ is_star = yes }}
					save_event_target_as = continuum_dsc_star
				}}
				spawn_megastructure = {{
					type = deep_space_citadel_0
					owner = event_target:{owner}
					planet = event_target:continuum_dsc_star
				}}
			}}""")
            continue
        if design:
            design_prep = f"""					event_target:{owner} = {{
						create_ship_design = {{ design = "{design}" }}
						add_ship_design = last_created_design
					}}
"""
            design_line = f'\n						design = "{design}"'
        else:
            design_prep = ""
            design_line = ""
        parts.append(f"""			every_system = {{
				limit = {{
					has_star_flag = {prefix}_{idx}_{short}
					NOT = {{ exists = space_owner }}
					NOT = {{ exists = starbase }}
				}}
				if = {{
					limit = {{ exists = event_target:{owner} }}
{design_prep}					create_starbase = {{
						size = {size}
						owner = event_target:{owner}{design_line}
					}}
				}}
			}}""")
    return "\n".join(parts)


def _rename_line(emp, pid, extra_indent=""):
    nm = (emp.get("colony_names") or {}).get(str(pid))
    if not nm:
        return ""
    return f"\n{extra_indent}						set_name = {script_name(nm)}"


def _pop_blocks(emp, prefix, idx, owner, species_tgt):
    pops = emp.get("colony_pop") or {}
    if not pops:
        return f"""			every_system = {{
				limit = {{ has_star_flag = {prefix}_{idx} }}
				every_system_planet = {{
					limit = {{
						OR = {{
							has_planet_flag = {prefix}_{idx}_homeworld
							has_planet_flag = {prefix}_{idx}_colony
						}}
						is_colonizable = yes
						NOT = {{ exists = owner }}
					}}
					if = {{
						limit = {{ exists = event_target:{owner} }}
						create_colony = {{
							owner = event_target:{owner}
							species = event_target:{species_tgt}
						}}
						create_pop_group = {{
							species = owner_main_species
							size = 1400
						}}
						remove_building = building_colony_shelter
					}}
				}}
			}}"""
    parts = []
    for i, (pid, n) in enumerate(pops.items()):
        rename = _rename_line(emp, pid)
        parts.append(f"""			every_system = {{
				limit = {{ has_star_flag = {prefix}_{idx} }}
				every_system_planet = {{
					limit = {{
						has_planet_flag = {prefix}_{idx}_c{i}
						is_colonizable = yes
						NOT = {{ exists = owner }}
					}}
					if = {{
						limit = {{ exists = event_target:{owner} }}
						create_colony = {{
							owner = event_target:{owner}
							species = event_target:{species_tgt}
						}}
						create_pop_group = {{
							species = owner_main_species
							size = {int(n)}
						}}{rename}
						remove_building = building_colony_shelter
					}}
				}}
			}}""")
    return "\n".join(parts)


def _fleet_block(emp, cap_flag, owner):
    mix = emp.get("fleet_mix") or {}
    fe_map = FE_SHIP_DESIGNS.get(_fe_key(emp), {}) if emp.get("type") == "fallen_empire" else {}
    bits = []
    for sz in FLEET_DESIGNS:
        if sz in _FLEET_BLOCK_SKIP:
            continue
        n = mix.get(sz)
        if not n:
            continue
        named = fe_map.get(sz)
        if named:
            ship = f'create_ship = {{ name = random design = "{named}" }}'
        else:
            ship = f"create_ship = {{ name = random random_existing_design = {sz} }}"
        bits.append(f"""								while = {{
									count = {int(n)}
									{ship}
								}}""")
    if not bits:
        if emp.get("type") == "fallen_empire":
            return ""
        bits.append("""								while = {
									count = 3
									create_ship = { name = random random_existing_design = corvette }
								}""")
    inner = "\n".join(bits)
    return f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {cap_flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							name = continuum_home_fleet
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
{inner}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}"""


def _fe_key(emp):
    auth = emp.get("authority") or ""
    civics = emp.get("civics") or []
    if auth == "auth_machine_intelligence" or any("machine" in (c or "") for c in civics):
        return "machine"
    ethics = [e for e in (emp.get("ethics") or []) if str(e).startswith("ethic_")]
    return emp.get("fallen_n") or FALLEN_BY_ETHIC.get(ethics[0] if ethics else "", "1")


def _fe_defense_block(emp, cap_flag, owner):
    """Spawn military_station_small_fallen_empire around FE capital. Count from emp fleet_mix or default 8."""
    mix = emp.get("fleet_mix") or {}
    n = mix.get("military_station_small_fallen_empire") or 8
    n = max(1, int(n))
    design = FE_PLATFORM_DESIGNS.get(_fe_key(emp)) or FE_PLATFORM_DESIGNS["1"]
    gfx = emp.get("graphical_culture") or "fallen_empire_01"
    return f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {cap_flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								while = {{
									count = {n}
									create_ship = {{ name = random design = "{design}" graphical_culture = {gfx} }}
								}}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}"""


def _fe_defense_platforms(emp, prefix, idx, owner):
    """Per-system FE defense platforms from Pre system_fleets. Fallback: capital mix."""
    design = FE_PLATFORM_DESIGNS.get(_fe_key(emp)) or FE_PLATFORM_DESIGNS["1"]
    gfx = emp.get("graphical_culture") or "fallen_empire_01"
    parts = []
    for sid, mix in (emp.get("system_fleets") or {}).items():
        n = int(mix.get("military_station_small_fallen_empire") or 0)
        if not n:
            continue
        flag = f"{prefix}_{idx}_s{sid}"
        parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								while = {{
									count = {n}
									create_ship = {{ name = random design = "{design}" graphical_culture = {gfx} }}
								}}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    if parts:
        return "\n".join(parts)
    return _fe_defense_block(emp, f"{prefix}_{idx}_capital", owner)


def _fe_mobile_fleets(emp, prefix, idx, owner, skip_sys=None):
    """Per-system FE warships from system_fleets, chunked at 40. Skip systems already restored as named fleets."""
    size_map = FE_SHIP_DESIGNS.get(_fe_key(emp), {}) or FE_SHIP_DESIGNS["1"]
    gfx = emp.get("graphical_culture") or "fallen_empire_01"
    skip = {str(s) for s in (skip_sys or []) if s}
    parts = []
    for sid, mix in (emp.get("system_fleets") or {}).items():
        if str(sid) in skip:
            continue
        flag = f"{prefix}_{idx}_s{sid}"
        for sz, design in size_map.items():
            n = int(mix.get(sz) or 0)
            if not n:
                continue
            left = n
            chunk_i = 0
            while left > 0:
                chunk = min(40, left)
                left -= chunk
                chunk_i += 1
                parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								while = {{
									count = {chunk}
									create_ship = {{ name = random design = "{design}" graphical_culture = {gfx} }}
								}}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    return "\n".join(parts)


def _angles(n):
    if n <= 0:
        return []
    return [int(i * 360 / n) for i in range(n)]


def _marauder_mix_for_system(emp, sid, is_capital):
    sys_mix = (emp.get("system_fleets") or {}).get(str(sid)) or {}
    if sys_mix:
        return sys_mix
    country = emp.get("fleet_mix") or {}
    if is_capital and any(country.get(sz) for sz in list(MARAUDER_SHIP_DESIGNS) + list(MARAUDER_STATION_DESIGNS)):
        return country
    return dict(MARAUDER_FLEET_DEFAULTS if is_capital else MARAUDER_SAT_DEFAULTS)


def _marauder_station_fleets(design, n, distance):
    parts = []
    for ang in _angles(int(n)):
        parts.append(f"""					create_fleet = {{
						settings = {{ spawn_debris = no garrison = yes }}
						effect = {{
							set_owner = prev
							create_ship = {{ name = random design = "{design}" prefix = no graphical_culture = pirate_01 }}
							set_location = {{ target = prevprev distance = {int(distance)} angle = {ang} }}
						}}
					}}""")
    return parts


def _marauder_garrison_fleet(mix, distance, angle):
    bits = []
    for sz, design in MARAUDER_SHIP_DESIGNS.items():
        if sz == "marauder_galleon":
            continue
        n = int(mix.get(sz) or 0)
        if not n:
            continue
        bits.append(f"""							while = {{
								count = {n}
								create_ship = {{ name = random design = "{design}" prefix = no graphical_culture = pirate_01 }}
							}}""")
    if not bits:
        return ""
    inner = "\n".join(bits)
    return f"""					create_fleet = {{
						settings = {{ spawn_debris = no garrison = yes }}
						effect = {{
							set_owner = prev
{inner}
							set_formation_scale = 2
							set_fleet_stance = aggressive
							set_aggro_range_measure_from = self
							set_aggro_range = 250
							set_location = {{ target = prevprev distance = {int(distance)} angle = {int(angle)} }}
						}}
					}}"""


def _marauder_galleon_fleets(n, distance):
    parts = []
    for ang in _angles(int(n)):
        parts.append(f"""					create_fleet = {{
						settings = {{ spawn_debris = no garrison = yes }}
						effect = {{
							set_owner = prev
							create_ship = {{ name = random design = "NAME_Ancestral_Glory" prefix = no graphical_culture = pirate_01 }}
							set_formation_scale = 2
							set_fleet_stance = aggressive
							set_aggro_range_measure_from = self
							set_aggro_range = 250
							set_location = {{ target = prevprev distance = {int(distance)} angle = {ang} }}
						}}
					}}""")
    return parts


def _marauder_system_contents(mix, is_capital):
    """Starbase + asteroid lairs + void dwellings + garrison fleets around the system star."""
    station_n = int(mix.get("marauder_station") or 0)
    dwelling_n = int(mix.get("marauder_void_dwelling") or 0)
    galleon_n = int(mix.get("marauder_galleon") or 0)
    station_dist = 230 if is_capital else 95
    dwelling_dist = 90 if is_capital else 65
    parts = []
    parts.extend(_marauder_station_fleets("NAME_Warrior_Freehold", station_n, station_dist))
    # Offset dwelling angles so they don't sit on top of asteroid lairs.
    dwell = _marauder_station_fleets("NAME_Void_Dwelling", dwelling_n, dwelling_dist)
    if not is_capital and dwelling_n >= 2:
        dwell = []
        for i, dist in enumerate((65, 115)[:dwelling_n]):
            ang = (0, 145, 220, 310)[i % 4]
            dwell.extend(_marauder_station_fleets("NAME_Void_Dwelling", 1, dist))
            dwell[-1] = dwell[-1].replace("angle = 0", f"angle = {ang}", 1)
    parts.extend(dwell)
    garrison = _marauder_garrison_fleet(mix, 100 if is_capital else 90, 120)
    if garrison:
        parts.append(garrison)
    parts.extend(_marauder_galleon_fleets(galleon_n, 20 if is_capital else 90))
    return "\n".join(parts)


def _marauder_presence(emp, idx, owner):
    """Per-system marauder starbases, stations, dwellings, and garrison fleets."""
    systems = emp.get("systems") or ([emp.get("home_system")] if emp.get("home_system") else [])
    systems = [str(s) for s in systems if s]
    if not systems:
        return ""
    chunks = []
    for k, sid in enumerate(systems):
        is_capital = k == 0
        mix = _marauder_mix_for_system(emp, sid, is_capital)
        flag = f"continuum_mar_{idx}_capital" if is_capital else f"continuum_mar_{idx}_n{k}"
        contents = _marauder_system_contents(mix, is_capital)
        contents_nl = ("\n" + contents) if contents else ""
        capital_line = ""
        if is_capital:
            capital_line = f"""
						event_target:{owner} = {{
							set_custom_capital_location = prev
						}}"""
        chunks.append(f"""			every_system = {{
				limit = {{
					has_star_flag = {flag}
					exists = event_target:{owner}
				}}
				if = {{
					limit = {{
						NOT = {{ exists = starbase }}
					}}
					create_starbase = {{
						size = starbase_marauder
						owner = event_target:{owner}
					}}
				}}
				every_system_planet = {{
					limit = {{ is_star = yes }}{capital_line}
					event_target:{owner} = {{
{contents_nl}
					}}
				}}
			}}""")
    return "\n".join(chunks)


def _inject_event_effects(event_txt, extra):
    if not extra or not extra.strip():
        return event_txt
    if not extra.endswith("\n"):
        extra += "\n"
    close = "		}\n	}\n}\n\n"
    i = event_txt.rfind(close)
    if i < 0:
        return event_txt + extra
    return event_txt[:i] + extra + event_txt[i:]


def write_intro_and_empire_events(events_dir, loc_entries, plan, empires, species, owned, capitals, trait_keys, civic_keys, ethic_keys, tech_keys=None):
    import os
    emp_path = os.path.join(events_dir, "continuum_empire_events.txt")
    chunks = []
    restored = []
    chunks.append("""namespace = continuum_empire
event = {
	id = continuum_empire.1
	is_triggered_only = yes
	hide_window = yes
	immediate = {
		if = {
			limit = { has_global_flag = continuum_empires_done }
		}
		else = {
			set_global_flag = continuum_empires_done
""")
    if plan.get("had_crisis"):
        chunks.append("			set_global_flag = continuum_had_crisis\n")
    chunks.append("""			random_playable_country = {
				limit = { is_ai = no }
				save_global_event_target_as = continuum_human
			}
""")
    for idx, emp in enumerate(empires):
        cid = emp["id"]
        syss = owned.get(cid) or []
        if not syss:
            continue
        sp = species.get(str(emp.get("founder_species"))) or {}
        flag = f"continuum_emp_{idx}"
        cap_flag = f"continuum_emp_{idx}_capital"
        name = script_name(emp.get("name") or "Unknown")
        sp_name = script_name(sp.get("name") or "Unknown")
        if is_machine(emp, sp):
            authority = "auth_machine_intelligence"
            civics = ["civic_machine_builder", "civic_machine_replication"]
            ethics = ["ethic_gestalt_consciousness"]
        elif is_hive(emp, sp):
            authority = "auth_hive_mind"
            civics = ["civic_hive_divided_attention", "civic_hive_one_mind"]
            ethics = ["ethic_gestalt_consciousness"]
        else:
            authority = emp.get("authority") or "auth_oligarchic"
            if authority in ("auth_machine_intelligence", "auth_hive_mind"):
                authority = "auth_oligarchic"
            ethics = [e for e in (emp.get("ethics") or []) if e != "ethic_gestalt_consciousness"]
            if ethic_keys:
                ethics = [e for e in ethics if e in ethic_keys]
            if not ethics:
                ethics = ["ethic_xenophile", "ethic_fanatic_materialist"]
            civics = [c for c in (emp.get("civics") or []) if c.startswith("civic_") and "machine" not in c and "hive" not in c]
            if civic_keys:
                filtered = [c for c in civics if c in civic_keys]
                if filtered:
                    civics = filtered
            if len(civics) < 2:
                civics = (civics + ["civic_mining_guilds", "civic_functional_architecture"])[:2]
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys]
        if is_machine(emp, sp):
            traits = [t for t in traits if "robot" in t or "machine" in t or "preference" in t] or ["trait_machine_unit"]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits[:8]) or "\t\t\t\ttrait = trait_adaptive"
        colors = emp.get("flag_colors") or ["red", "black", "black", "null"]
        while len(colors) < 4:
            colors.append("null")
        gfx = emp.get("graphical_culture") or "mammalian_01"
        tech_lines = _tech_effect_lines(emp, tech_keys)
        prefix_line = ""
        resolved_pre = continuum_cp.resolved_ship_prefix(emp)
        if resolved_pre:
            emp["ship_prefix"] = resolved_pre
            prefix_line = f'\n					set_ship_prefix = "{resolved_pre}"'
        restored.append({
            "emp": emp,
            "prefix": "continuum_emp",
            "idx": idx,
            "owner": f"continuum_emp_{idx}",
            "cap_flag": cap_flag,
            "_sp_class": (species.get(str(emp.get("founder_species"))) or {}).get("class") or "",
            "pop_txt": _pop_blocks(emp, "continuum_emp", idx, f"continuum_emp_{idx}", f"continuum_sp_{idx}"),
            "sb_txt": _starbase_blocks("continuum_emp", idx, f"continuum_emp_{idx}", emp),
        })
        chunks.append(f"""			create_species = {{
				name = {sp_name}
				class = {sp.get('class', 'MAM')}
				portrait = {sp.get('portrait', 'human')}
				namelist = {sp.get('namelist', 'MAM1')}
				traits = {{
					ideal_planet_class = pc_continental
{trait_lines}
				}}
			}}
			last_created_species = {{
				save_global_event_target_as = continuum_sp_{idx}
			}}
			create_country = {{
				name = {name}
				type = default
				authority = {authority}
				civics = {{ civic = {civics[0]} civic = {civics[1]} }}
				origin = origin_default
				species = last_created_species
				ethos = {{ {' '.join(f'ethic = {e}' for e in ethics[:3])} }}
				flag = {{
					icon = {{ category = "{emp.get('flag_icon_cat', 'special')}" file = "{emp.get('flag_icon', 'pirate_flag.dds')}" }}
					background = {{ category = "{emp.get('flag_bg_cat', 'backgrounds')}" file = "{emp.get('flag_bg', '00_solid.dds')}" }}
					colors = {{ "{colors[0]}" "{colors[1]}" "{colors[2]}" "{colors[3]}" }}
				}}
				ignore_initial_colony_error = yes
				day_zero_contact = no
				exclude_day_zero_contact = event_target:continuum_human
				effect = {{
					save_global_event_target_as = continuum_emp_{idx}
					set_graphical_culture = {gfx}
					set_country_flag = continuum_pre_empire{prefix_line}
{tech_lines}
					add_resource = {{ energy = 1000 minerals = 1000 food = 1000 alloys = 500 influence = 200 }}
				}}
			}}
""")

    # Enclave names from Pre (vanilla Coven/Salvager roll a random name).
    enc = plan.get("enclave_names") or {}
    if enc.get("shroudwalker"):
        nm = script_name(enc["shroudwalker"])
        chunks.append(f"""			every_country = {{
				limit = {{ has_country_flag = shroudwalker_enclave_country }}
				set_name = {nm}
			}}
""")
    if enc.get("salvager"):
        nm = script_name(enc["salvager"])
        chunks.append(f"""			every_country = {{
				limit = {{ has_country_flag = salvager_enclave_country }}
				set_name = {nm}
			}}
""")

    fallen = plan.get("fallen") or []
    fe_owned = plan.get("fallen_owned") or {}
    for idx, emp in enumerate(fallen):
        if not fe_owned.get(emp["id"]):
            continue
        sp = species.get(str(emp.get("founder_species"))) or {}
        name = script_name(emp.get("name") or "Unknown")
        fe_loc = f"continuum_fe_name_{idx}"
        loc_entries[fe_loc] = display_empire_name(emp.get("name") or "Unknown")
        sp_name = script_name(sp.get("name") or "Unknown")
        ethics = [e for e in (emp.get("ethics") or []) if e.startswith("ethic_")]
        if ethic_keys:
            ethics = [e for e in ethics if e in ethic_keys]
        if not ethics:
            ethics = ["ethic_fanatic_materialist"]
        fe_n = emp.get("fallen_n") or FALLEN_BY_ETHIC.get(ethics[0], "1")
        gfx = emp.get("graphical_culture") or f"fallen_empire_0{fe_n}"
        designs = FALLEN_DESIGNS.get(fe_n) or FALLEN_DESIGNS["1"]
        size_map = dict(FE_SHIP_DESIGNS.get(fe_n) or FE_SHIP_DESIGNS["1"])
        plat = FE_PLATFORM_DESIGNS.get(fe_n)
        if plat:
            size_map["military_station_small_fallen_empire"] = plat
        emp["_fe_size_designs"] = size_map
        design_lines = "\n".join(f"					add_global_ship_design = \"{d}\"" for d in designs)
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys][:8]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits) or "\t\t\t\ttrait = trait_adaptive"
        colors = emp.get("flag_colors") or ["black", "black", "black", "null"]
        while len(colors) < 4:
            colors.append("null")
        tech_lines = _tech_effect_lines(emp, tech_keys)
        restored.append({
            "emp": emp,
            "prefix": "continuum_fe",
            "idx": idx,
            "owner": f"continuum_fe_{idx}",
            "cap_flag": f"continuum_fe_{idx}_capital",
            "pop_txt": _pop_blocks(emp, "continuum_fe", idx, f"continuum_fe_{idx}", f"continuum_fe_sp_{idx}"),
            "sb_txt": _starbase_blocks("continuum_fe", idx, f"continuum_fe_{idx}", emp),
        })
        chunks.append(f"""			create_species = {{
				name = {sp_name}
				class = {sp.get('class', 'MAM')}
				portrait = {sp.get('portrait', 'human')}
				namelist = {sp.get('namelist', 'MAM1')}
				traits = {{
					ideal_planet_class = pc_continental
{trait_lines}
				}}
			}}
			last_created_species = {{
				save_global_event_target_as = continuum_fe_sp_{idx}
			}}
			create_country = {{
				name = {fe_loc}
				type = fallen_empire
				authority = auth_imperial
				civics = {{ civic = civic_lethargic_leadership civic = civic_empire_in_decline }}
				origin = origin_fallen_empire
				species = last_created_species
				ethos = {{ {' '.join(f'ethic = {e}' for e in ethics[:3])} }}
				flag = {{
					icon = {{ category = "{emp.get('flag_icon_cat', 'special')}" file = "{emp.get('flag_icon', 'pirate_flag.dds')}" }}
					background = {{ category = "{emp.get('flag_bg_cat', 'backgrounds')}" file = "{emp.get('flag_bg', '00_solid.dds')}" }}
					colors = {{ "{colors[0]}" "{colors[1]}" "{colors[2]}" "{colors[3]}" }}
				}}
				ignore_initial_colony_error = yes
				day_zero_contact = no
				exclude_day_zero_contact = event_target:continuum_human
				effect = {{
					save_global_event_target_as = continuum_fe_{idx}
					set_graphical_culture = {gfx}
					set_country_flag = fallen_empire_{fe_n}
					set_country_flag = continuum_pre_fallen
					set_name = {fe_loc}
					add_resource = {{ minerals = 10000 energy = 10000 food = 1000 influence = 500 }}
{design_lines}
{tech_lines}
				}}
			}}
			last_created_country = {{
				set_name = {fe_loc}
			}}
""")

    for idx, mar in enumerate(plan.get("marauders") or []):
        if not mar.get("home_system"):
            continue
        sp = species.get(str(mar.get("founder_species"))) or {}
        name = script_name(mar.get("name") or "Unknown")
        mar_loc = f"continuum_mar_name_{idx}"
        loc_entries[mar_loc] = display_empire_name(mar.get("name") or "Unknown")
        sp_name = script_name(sp.get("name") or "Unknown")
        n = mar.get("marauder_n") or str(idx + 1)
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys][:8]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits) or "\t\t\t\ttrait = trait_rapid_breeders"
        colors = mar.get("flag_colors") or ["black", "black", "null", "null"]
        while len(colors) < 4:
            colors.append("null")
        icon_cat = mar.get("flag_icon_cat") or "pirate"
        icon = mar.get("flag_icon") or "flag_pirate_7.dds"
        bg_cat = mar.get("flag_bg_cat") or "backgrounds"
        bg = mar.get("flag_bg") or "00_solid.dds"
        design_lines = "\n".join(f'						add_global_ship_design = "{d}"' for d in MARAUDER_GLOBAL_DESIGNS)
        presence = _marauder_presence(mar, idx, f"continuum_mar_{idx}")
        chunks.append(f"""			every_system = {{
				limit = {{ has_star_flag = continuum_mar_{idx}_capital }}
				create_species = {{
					name = {sp_name}
					class = {sp.get('class', 'MAM')}
					portrait = {sp.get('portrait', 'human')}
					namelist = {sp.get('namelist', 'MAM1')}
					traits = {{
						ideal_planet_class = pc_habitat
{trait_lines}
					}}
				}}
				create_country = {{
					name_list = {sp.get('namelist', 'MAM1')}
					type = dormant_marauders
					civics = {{ civic = civic_anarcho_tribalism }}
					origin = origin_default
					species = last_created_species
					ethos = {{ ethic = ethic_fanatic_militarist ethic = ethic_xenophobe }}
					flag = {{
						icon = {{ category = "{icon_cat}" file = "{icon}" }}
						background = {{ category = "{bg_cat}" file = "{bg}" }}
						colors = {{ "{colors[0]}" "{colors[1]}" "{colors[2]}" "{colors[3]}" }}
					}}
					ignore_initial_colony_error = yes
					day_zero_contact = no
					exclude_day_zero_contact = event_target:continuum_human
					effect = {{
						save_global_event_target_as = continuum_mar_{idx}
						set_graphical_culture = pirate_01
						set_country_flag = marauder_{n}
						set_country_flag = continuum_pre_marauder
						set_name = {mar_loc}
						create_ship_design = {{ design = "NAME_Marauder_Starbase" }}
						add_ship_design = last_created_design
{design_lines}
					}}
				}}
			}}
{presence}
""")

    for idx, prim in enumerate(plan.get("primitives") or []):
        sp = species.get(str(prim.get("founder_species"))) or {}
        name = script_name(prim.get("name") or "Unknown")
        sp_name = script_name(sp.get("name") or "Unknown")
        ethics = [e for e in (prim.get("ethics") or []) if e.startswith("ethic_")]
        if ethic_keys:
            ethics = [e for e in ethics if e in ethic_keys]
        if not ethics:
            ethics = ["ethic_xenophile", "ethic_fanatic_egalitarian"]
        civics = [c for c in (prim.get("civics") or []) if c.startswith("civic_")]
        if len(civics) < 2:
            civics = ["civic_secret_of_fire", "civic_the_wheel"]
        gfx = prim.get("graphical_culture") or "preindustrial_01"
        age = prim.get("pre_ftl_age")
        if not age:
            age = "industrial_age" if "industrial" in gfx else "iron_age"
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys][:8]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits) or "\t\t\t\ttrait = trait_adaptive"
        origin = prim.get("origin") or "origin_default_pre_ftl"
        if (
            not str(origin).startswith("origin_")
            or "pre_ftl" not in str(origin)
        ):
            origin = "origin_default_pre_ftl"
        colors = prim.get("flag_colors") or ["turquoise", "green", "null", "null"]
        while len(colors) < 4:
            colors.append("null")
        prim_rename = ""
        pnames = prim.get("colony_names") or {}
        if pnames:
            prim_rename = f"\n\t\t\t\tset_name = {script_name(next(iter(pnames.values())))}"
        n_prim_army = sum(1 for a in (prim.get("armies") or []) if (a.get("type") or "") == "primitive_army")
        army_txt = ""
        if n_prim_army:
            army_txt = "\n				every_planet_army = { remove_army = yes }\n" + "\n".join(
                f"				create_army = {{ owner = event_target:continuum_prim_{idx} species = last_created_species type = primitive_army }}"
                for _ in range(int(n_prim_army))
            )
        chunks.append(f"""			every_system = {{
				limit = {{ has_star_flag = continuum_prim_{idx} }}
				every_system_planet = {{
				limit = {{
					has_planet_flag = continuum_prim_{idx}_homeworld
					is_colony = no
					NOT = {{ exists = owner }}
				}}
				create_species = {{
					name = {sp_name}
					class = {sp.get('class', 'MAM')}
					portrait = {sp.get('portrait', 'human')}
					namelist = {sp.get('namelist', 'MAM1')}
					homeworld = this
					traits = {{
						ideal_planet_class = pc_continental
{trait_lines}
					}}
				}}
				create_country = {{
					name = {name}
					type = primitive
					authority = {prim.get('authority') or 'auth_oligarchic'}
					civics = {{ civic = {civics[0]} civic = {civics[1]} }}
					origin = {origin}
					species = last_created_species
					ethos = {{ {' '.join(f'ethic = {e}' for e in ethics[:3])} }}
					flag = {{
						icon = {{ category = "{prim.get('flag_icon_cat', 'pre_ftl')}" file = "{prim.get('flag_icon', 'preftl_stone_age.dds')}" }}
						background = {{ category = "{prim.get('flag_bg_cat', 'backgrounds')}" file = "{prim.get('flag_bg', 'new_dawn.dds')}" }}
						colors = {{ "{colors[0]}" "{colors[1]}" "{colors[2]}" "{colors[3]}" }}
					}}
					day_zero_contact = no
					ignore_initial_colony_error = yes
					effect = {{
						set_graphical_culture = {gfx}
						set_country_flag = {age}
						set_pre_ftl_age = {age}
						set_country_flag = continuum_pre_primitive
						save_global_event_target_as = continuum_prim_{idx}
					}}
				}}
				create_colony = {{
					owner = last_created_country
					species = last_created_species
				}}
				create_pop_group = {{
					species = owner_main_species
					size = {int(next(iter((prim.get("colony_pop") or {}).values()), 1400))}
				}}{prim_rename}{army_txt}
				}}
			}}
""")

    import continuum_player
    chunks.append(continuum_player.emit_copy_identity_effects(restored, loc_entries))
    for idx, emp in enumerate(plan.get("fallen") or []):
        if not (plan.get("fallen_owned") or {}).get(emp["id"]):
            continue
        chunks.append(f"""			if = {{
				limit = {{ exists = event_target:continuum_fe_{idx} }}
				event_target:continuum_fe_{idx} = {{ set_name = continuum_fe_name_{idx} }}
			}}
""")
    for idx, mar in enumerate(plan.get("marauders") or []):
        if not mar.get("home_system"):
            continue
        chunks.append(f"""			if = {{
				limit = {{ exists = event_target:continuum_mar_{idx} }}
				event_target:continuum_mar_{idx} = {{ set_name = continuum_mar_name_{idx} }}
			}}
""")
    chunks.append("		}\n	}\n}\n\n")
    chunks.append(continuum_cp.emit_pop_event(restored))
    chunks.append(continuum_cp.emit_starbase_event(restored))
    fleet_txt = continuum_cp.emit_fleet_event(restored, _fleet_block)
    fe_bits = []
    for r in restored:
        if r["emp"].get("type") == "fallen_empire" or str(r.get("prefix") or "").startswith("continuum_fe"):
            fe_bits.append(_fe_defense_platforms(r["emp"], r["prefix"], r["idx"], r["owner"]))
            named_sys = {str(fl.get("sys")) for fl in (r["emp"].get("named_military") or []) if fl.get("sys")}
            fe_bits.append(_fe_mobile_fleets(r["emp"], r["prefix"], r["idx"], r["owner"], skip_sys=named_sys))
    fe_def = "\n".join(p for p in fe_bits if p)
    later_names = []
    for idx, emp in enumerate(plan.get("fallen") or []):
        if (plan.get("fallen_owned") or {}).get(emp["id"]):
            later_names.append(f"""			if = {{
				limit = {{ exists = event_target:continuum_fe_{idx} }}
				event_target:continuum_fe_{idx} = {{ set_name = continuum_fe_name_{idx} }}
			}}""")
    for idx, mar in enumerate(plan.get("marauders") or []):
        if mar.get("home_system"):
            later_names.append(f"""			if = {{
				limit = {{ exists = event_target:continuum_mar_{idx} }}
				event_target:continuum_mar_{idx} = {{ set_name = continuum_mar_name_{idx} }}
			}}""")
    extra = "\n".join(p for p in (fe_def, "\n".join(later_names)) if p)
    chunks.append(_inject_event_effects(fleet_txt, extra))
    chunks.append(continuum_cp.emit_station_event(restored))
    chunks.append(continuum_cp.emit_module_mega_event(restored))
    chunks.append(continuum_cp.emit_leader_event(restored))
    chunks.append(continuum_cp.emit_diplo_event(restored))
    chunks.append(continuum_cp.emit_intel_event(restored))
    prim_restored = []
    for idx, prim in enumerate(plan.get("primitives") or []):
        if not prim.get("system_id"):
            continue
        prim_restored.append({
            "emp": prim,
            "prefix": "continuum_prim",
            "idx": idx,
            "owner": f"continuum_prim_{idx}",
            "cap_flag": f"continuum_prim_{idx}",
            "pop_txt": "",
            "sb_txt": "",
        })
    layout_restored = restored + prim_restored
    chunks.append(continuum_cp.emit_layout_event(layout_restored))
    chunks.append(continuum_cp.emit_army_event(layout_restored, include_embarked=True))
    chunks.append(continuum_cp.emit_site_event(restored))
    chunks.append(continuum_cp.emit_tradition_event(restored))
    with open(emp_path, "w", encoding="utf-8") as f:
        f.write("".join(chunks))
    player_path = os.path.join(events_dir, "continuum_player_events.txt")
    with open(player_path, "w", encoding="utf-8") as f:
        f.write(continuum_player.emit_player_event(restored, loc_entries))
    intro_path = os.path.join(events_dir, "continuum_intro_events.txt")
    with open(intro_path, "w", encoding="utf-8") as f:
        f.write(continuum_player.emit_intro_event(restored, loc_entries, plan.get("had_crisis")))

    loc_entries["continuum_home_fleet"] = "Home Fleet"
    loc_entries["opinion_continuum_ceasefire"] = "Lingering Hostility"
    # Vanilla 4.4 GUI uses this key; english loc never defines it.
    loc_entries["OPINION_LEVEL"] = "Opinion"


def _tag_owned_planets(countries, colony_to_planet, prefix, flags, planet_flags, owned, capitals):
    for idx, emp in enumerate(countries):
        for sid in owned.get(emp["id"], []):
            flags.setdefault(sid, []).append(f"{prefix}_{idx}")
            fl = flags.get(str(sid), [])
            if "continuum_border" in fl or "continuum_front" in fl:
                flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_border")
        cap = capitals.get(emp["id"])
        if cap:
            flags.setdefault(cap, []).append(f"{prefix}_{idx}_capital")
        for token in emp.get("owned_planets") or []:
            pid = colony_to_planet.get(str(token))
            if pid:
                planet_flags.setdefault(str(pid), []).append(f"{prefix}_{idx}_colony")
        cap_col = emp.get("capital") or ((emp.get("owned_planets") or [None])[0])
        cap_pid = colony_to_planet.get(str(cap_col)) if cap_col is not None else None
        if cap_pid:
            planet_flags.setdefault(str(cap_pid), []).append(f"{prefix}_{idx}_homeworld")
        for sid, size in (emp.get("starbases") or {}).items():
            short = STARBASE_SHORT.get(size)
            if short:
                flags.setdefault(str(sid), []).append(f"{prefix}_{idx}")
                flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_{short}")
        for i, pid in enumerate((emp.get("colony_pop") or {}).keys()):
            planet_flags.setdefault(str(pid), []).append(f"{prefix}_{idx}_c{i}")
        continuum_cp.tag_cp_flags([emp], prefix, flags, planet_flags, {emp["id"]: idx})


def _merge_starbase_systems(countries, owned, types):
    claimed = {s for cid, syss in owned.items() for s in syss}
    for c in countries:
        if c["type"] not in types:
            continue
        syss = owned.setdefault(c["id"], [])
        mine = set(syss)
        cleaned = {}
        for sid, size in (c.get("starbases") or {}).items():
            if sid in claimed and sid not in mine:
                continue
            cleaned[sid] = size
            if sid not in syss:
                syss.append(sid)
                claimed.add(sid)
        for sid in syss:
            cleaned.setdefault(sid, "starbase_starport" if c["type"] == "default" else "starbase_citadel")
        c["starbases"] = cleaned


def build_empire_plan(save_path, galaxy_data, stars, planets, argv_start=None):
    countries, species, had_crisis, _data, colony_to_planet = load_empire_save_data(save_path)
    p2s = planet_to_system_map(stars)
    attach_primitive_systems(countries, p2s, colony_to_planet)
    attach_snapshot(_data, countries, colony_to_planet, p2s)
    attach_colony_names(countries, names_from_galaxy(galaxy_data), colony_to_planet)
    habitables = system_habitables(galaxy_data, planets)
    blocked = blocked_systems(galaxy_data)
    fe_owned, fe_caps = ownership_from_planets(countries, p2s, colony_to_planet, types=("fallen_empire",))
    _merge_starbase_systems(countries, fe_owned, ("fallen_empire",))
    for syss in fe_owned.values():
        blocked.update(syss)
    marauders = [c for c in countries if c["type"] == "dormant_marauders"]
    marauder_homes = {}
    marauder_sys_ids = []
    for sys in galaxy_data:
        sid = str(sys.get("id"))
        for fl in sys.get("flags") or []:
            sfl = str(fl)
            if sfl.startswith("marauder_capital_"):
                marauder_homes[sfl] = sid
                blocked.add(sid)
            if sfl in ("marauder_system", "marauder_capital_1", "marauder_capital_2", "marauder_capital_3"):
                blocked.add(sid)
            if sfl == "marauder_system" and sid not in marauder_sys_ids:
                marauder_sys_ids.append(sid)
    owned, capitals = ownership_from_planets(countries, p2s, colony_to_planet, types=("default",))
    _merge_starbase_systems(countries, owned, ("default",))
    for cid, syss in list(owned.items()):
        cap = capitals.get(cid)
        owned[cid] = [s for s in syss if s not in blocked or s == cap]
    lanes = hyperlane_map(galaxy_data)
    tags, spawn_weights = classify_systems(owned, capitals, lanes, blocked, habitables, countries, species)
    defaults = [c for c in countries if c["type"] == "default" and owned.get(c["id"])]
    fallens = [c for c in countries if c["type"] == "fallen_empire" and fe_owned.get(c["id"])]
    primitives = [c for c in countries if c["type"] == "primitive" and c.get("system_id")]
    if defaults:
        te = defaults[0]
        print(
            f"[CP] {te.get('name')} id={te.get('id')} "
            f"systems={owned.get(te['id'])} "
            f"pops={list((te.get('colony_pop') or {}).values())} "
            f"names={list((te.get('colony_names') or {}).values())} "
            f"{continuum_cp.cp_summary(te)} "
            f"starbases={len(te.get('starbases') or {})}"
        )
    flags = {}
    planet_flags = {}
    for sid, flist in tags.items():
        flags.setdefault(sid, []).extend(flist)
    _tag_owned_planets(defaults, colony_to_planet, "continuum_emp", flags, planet_flags, owned, capitals)
    _tag_owned_planets(fallens, colony_to_planet, "continuum_fe", flags, planet_flags, fe_owned, fe_caps)
    for fe in fallens:
        plat = (fe.get("fleet_mix") or {}).get("military_station_small_fallen_empire")
        print(
            f"[CP] fallen {fe.get('name')} systems={fe_owned.get(fe['id'])} "
            f"starbases={fe.get('starbases')} platforms={plat} per_sys={fe.get('system_fleets')}"
        )
    for idx, prim in enumerate(primitives):
        sid = prim.get("system_id")
        if sid:
            flags.setdefault(sid, []).append(f"continuum_prim_{idx}")
        cap_col = prim.get("capital") or ((prim.get("owned_planets") or [None])[0])
        cap_pid = colony_to_planet.get(str(cap_col)) if cap_col is not None else None
        if cap_pid:
            planet_flags.setdefault(str(cap_pid), []).append(f"continuum_prim_{idx}_homeworld")
        for i, pid in enumerate((prim.get("colony_pop") or {}).keys()):
            planet_flags.setdefault(str(pid), []).append(f"continuum_prim_{idx}_c{i}")
        if cap_pid:
            planet_flags.setdefault(str(cap_pid), []).append(f"continuum_prim_{idx}_c0")
    used_mar_sys = set()
    for idx, mar in enumerate(marauders):
        n = mar.get("marauder_n") or str(idx + 1)
        cap_sid = marauder_homes.get(f"marauder_capital_{n}") or marauder_homes.get(f"marauder_capital_{idx + 1}")
        systems = []
        if cap_sid:
            systems.append(cap_sid)
            used_mar_sys.add(cap_sid)
        q = [cap_sid] if cap_sid else []
        seen = set(q)
        qi = 0
        while qi < len(q) and len(systems) < 3:
            cur = q[qi]
            qi += 1
            for nb in lanes.get(cur) or []:
                nb = str(nb)
                if nb in seen:
                    continue
                seen.add(nb)
                q.append(nb)
                if nb in marauder_sys_ids and nb not in used_mar_sys:
                    systems.append(nb)
                    used_mar_sys.add(nb)
                    if len(systems) >= 3:
                        break
        if len(systems) < 3:
            for sid in marauder_sys_ids:
                if sid not in used_mar_sys:
                    systems.append(sid)
                    used_mar_sys.add(sid)
                    if len(systems) >= 3:
                        break
        mar["marauder_n"] = n
        mar["systems"] = systems
        print(f"[CP] marauder {mar.get('name')} n={n} systems={systems} mix={mar.get('fleet_mix')} per_sys={mar.get('system_fleets')}")
        if systems:
            mar["home_system"] = systems[0]
            for k, sid in enumerate(systems):
                flags.setdefault(sid, []).append(f"continuum_mar_{idx}")
                if k == 0:
                    flags.setdefault(sid, []).append(f"continuum_mar_{idx}_capital")
                else:
                    flags.setdefault(sid, []).append(f"continuum_mar_{idx}_n{k}")
    enclave_names = {}
    for c in countries:
        if c["type"] != "enclave":
            continue
        civics = set(c.get("civics") or [])
        if "civic_shroudwalker_enclave" in civics:
            enclave_names["shroudwalker"] = c.get("name")
        elif "civic_salvager_enclave" in civics:
            enclave_names["salvager"] = c.get("name")
    apply_home_prime(galaxy_data, defaults, capitals)
    spawn_ids = set(spawn_weights.keys())
    fallback = next(iter(spawn_ids), str(galaxy_data[0].get("id")) if galaxy_data else "0")
    return {
        "kind": "auto",
        "player_system": fallback,
        "had_crisis": had_crisis,
        "owned": owned,
        "capitals": capitals,
        "empires": defaults,
        "fallen": fallens,
        "fallen_owned": fe_owned,
        "fallen_capitals": fe_caps,
        "marauders": marauders,
        "primitives": primitives,
        "enclave_names": enclave_names,
        "species": species,
        "spawn_ids": spawn_ids,
        "spawn_weights": spawn_weights,
        "extra_flags": flags,
        "planet_flags": planet_flags,
        "countries": countries,
        "devastation_system": None,
    }
