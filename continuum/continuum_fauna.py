"""Snapshot Pre fauna / unique ships per system and emit Continuum Present spawns."""
import re
import zipfile
from collections import defaultdict

FAUNA_DESIGNS = {
    "crystal_ship_small_blue": ("NAME_Small_Crystal_Entity_Blue", "crystal", "NAME_Sapphire_Lurkers"),
    "crystal_ship_medium_blue": ("NAME_Medium_Crystal_Entity_Blue", "crystal", "NAME_Sapphire_Lurkers"),
    "crystal_ship_large_blue": ("NAME_Large_Crystal_Entity_Blue", "crystal", "NAME_Sapphire_Lurkers"),
    "crystal_ship_small_green": ("NAME_Small_Crystal_Entity_Green", "crystal", "NAME_Emerald_Roamers"),
    "crystal_ship_medium_green": ("NAME_Medium_Crystal_Entity_Green", "crystal", "NAME_Emerald_Roamers"),
    "crystal_ship_large_green": ("NAME_Large_Crystal_Entity_Green", "crystal", "NAME_Emerald_Roamers"),
    "crystal_ship_small_red": ("NAME_Small_Crystal_Entity_Red", "crystal", "NAME_Ruby_Stack"),
    "crystal_ship_medium_red": ("NAME_Medium_Crystal_Entity_Red", "crystal", "NAME_Ruby_Stack"),
    "crystal_ship_large_red": ("NAME_Large_Crystal_Entity_Red", "crystal", "NAME_Ruby_Stack"),
    "crystal_ship_small_yellow": ("NAME_Small_Crystal_Entity_Yellow", "crystal", "NAME_Topaz_Guardians"),
    "crystal_ship_medium_yellow": ("NAME_Medium_Crystal_Entity_Yellow", "crystal", "NAME_Topaz_Guardians"),
    "crystal_ship_large_yellow": ("NAME_Large_Crystal_Entity_Yellow", "crystal", "NAME_Topaz_Guardians"),
    "crystal_ship_small_blue_elite": ("NAME_Small_Crystal_Entity_Blue_Elite", "crystal", "NAME_Sapphire_Guardians"),
    "crystal_ship_medium_blue_elite": ("NAME_Medium_Crystal_Entity_Blue_Elite", "crystal", "NAME_Sapphire_Guardians"),
    "crystal_ship_large_blue_elite": ("NAME_Large_Crystal_Entity_Blue_Elite", "crystal", "NAME_Sapphire_Guardians"),
    "crystal_ship_small_green_elite": ("NAME_Small_Crystal_Entity_Green_Elite", "crystal", "NAME_Emerald_Guardians"),
    "crystal_ship_medium_green_elite": ("NAME_Medium_Crystal_Entity_Green_Elite", "crystal", "NAME_Emerald_Guardians"),
    "crystal_ship_large_green_elite": ("NAME_Large_Crystal_Entity_Green_Elite", "crystal", "NAME_Emerald_Guardians"),
    "crystal_ship_small_red_elite": ("NAME_Small_Crystal_Entity_Red_Elite", "crystal", "NAME_Ruby_Guardians"),
    "crystal_ship_medium_red_elite": ("NAME_Medium_Crystal_Entity_Red_Elite", "crystal", "NAME_Ruby_Guardians"),
    "crystal_ship_large_red_elite": ("NAME_Large_Crystal_Entity_Red_Elite", "crystal", "NAME_Ruby_Guardians"),
    "crystal_ship_small_yellow_elite": ("NAME_Small_Crystal_Entity_Yellow_Elite", "crystal", "NAME_Topaz_Guardians"),
    "crystal_ship_medium_yellow_elite": ("NAME_Medium_Crystal_Entity_Yellow_Elite", "crystal", "NAME_Topaz_Guardians"),
    "crystal_ship_large_yellow_elite": ("NAME_Large_Crystal_Entity_Yellow_Elite", "crystal", "NAME_Topaz_Guardians"),
    "crystal_station_large": ("NAME_Crystal_Nidus", "crystal", "NAME_Crystal_Nidus"),
    "space_whale_1": ("NAME_Tiyanki_Cow", "tiyanki", "NAME_Tiyanki_Space_Whale"),
    "space_whale_2": ("NAME_Tiyanki_Bull", "tiyanki", "NAME_Tiyanki_Space_Whale"),
    "space_whale_3": ("NAME_Tiyanki_Hatchling", "tiyanki", "NAME_Tiyanki_Space_Whale"),
    "space_whale_4": ("NAME_Tiyanki_Calf", "tiyanki", "NAME_Tiyanki_Space_Whale"),
    "space_whale_5": ("NAME_Tiyanki_Ox", "tiyanki", "NAME_Tiyanki_Space_Whale"),
    "space_amoeba": ("NAME_Small_Space_Organism_Zebra", "amoeba", "NAME_Space_Amoeba_plural"),
    "space_amoeba_mother": ("NAME_Large_Space_Organism_Zebra", "amoeba", "NAME_Space_Amoeba_plural"),
    "ancient_mining_drone": ("NAME_Ancient_Mining_Drone", "drone", "NAME_Ancient_Mining_Drones"),
    "ancient_corvette": ("NAME_Ancient_Combat_Drone", "drone", "NAME_Ancient_Mining_Drones"),
    "ancient_destroyer": ("NAME_Ancient_Destroyer", "drone", "NAME_Asset_Protection_Unit"),
    "space_cloud": ("NAME_Cloud_Entity", "cloud", "NAME_Void_Cloud"),
    "voidworms_small": ("NAME_Voidworms_Nymph", "voidworms", "NAME_Voidworms"),
    "voidworms_medium": ("NAME_Voidworms_Juvenile", "voidworms", "NAME_Voidworms"),
    "voidworms_large": ("NAME_Voidworms_Mature", "voidworms", "NAME_Voidworms"),
    "voidworms_titan": ("NAME_Voidworms_Troika", "voidworms", "NAME_Voidworms"),
    "voidworm_nest": ("NAME_Voidworms_Starbase", "voidworms", "NAME_Voidworms_Starbase"),
    "wenkwort_drone": ("NAME_Gardener_Drone", "gardener", "NAME_Gardener_Drone_plural"),
    "ghost_ship": ("NAME_Hillos_ship", "hillos", "NAME_Hillos_ship"),
    "lost_swarm_adult": ("NAME_Lost_Swarm_Adult", "lost_swarm", "NAME_Lost_Swarm"),
    "psionic_avatar": ("NAME_Shroud_Avatar", "coven", "NAME_Psionic_Avatar"),
}

GARRISON_CTYPES = {"tiyanki_garrison": "tiyanki", "amoeba_garrison": "amoeba"}


def _brace(data, key):
    needle = f"\n{key}=\n{{"
    i = data.find(needle)
    if i < 0:
        return ""
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


def _split(section, tab="\t"):
    out = {}
    hdr = re.compile(rf"\n{tab}(\d+)=\n{tab}\{{")
    i = 0
    n = len(section)
    while True:
        m = hdr.search(section, i)
        if not m:
            break
        start = section.find("{", m.start())
        depth = 0
        j = start
        while j < n:
            if section[j] == "{":
                depth += 1
            elif section[j] == "}":
                depth -= 1
                if depth == 0:
                    out[m.group(1)] = section[start + 1:j]
                    i = j + 1
                    break
            j += 1
        else:
            break
    return out


def parse_fauna_snapshot(save_path):
    with zipfile.ZipFile(save_path, "r") as z:
        data = z.read("gamestate").decode("utf-8", "replace")
    ctypes = {}
    fleet_owner = {}
    for cid, body in _split(_brace(data, "country")).items():
        tm = re.search(r'\n\t\ttype="([^"]+)"', body)
        ctypes[cid] = tm.group(1) if tm else ""
        start = body.find("owned_fleets=")
        chunk = body[start:] if start >= 0 else ""
        for m in re.finditer(r"fleet=(\d+)", chunk):
            fleet_owner.setdefault(m.group(1), cid)
        fm = re.search(r"\nfleets=\s*\n\t\t\{([^}]+)\}", body)
        if fm:
            for fid in fm.group(1).split():
                if fid.isdigit():
                    fleet_owner.setdefault(fid, cid)
    designs = {}
    for did, body in _split(_brace(data, "ship_design")).items():
        m = re.search(r'ship_size="([^"]+)"', body)
        if m:
            designs[did] = m.group(1)
    fleet_names = {}
    for fid, body in _split(_brace(data, "fleet")).items():
        keys = re.findall(r'key="([^"]+)"', body[:900])
        name = None
        for k in keys:
            if k.startswith("NAME_"):
                name = k
                break
        if not name:
            for k in keys:
                if k not in ("NAME", "PREFIX", "SUFFIX", "1", "2", "3") and not k.startswith("%"):
                    name = k
                    break
        fleet_names[fid] = name
    fleets_out = {}
    drone_mines = defaultdict(int)
    drone_mine_planets = []
    systems = set()
    try:
        import continuum_cp
        planet_xy = continuum_cp._planet_xy(data)
    except Exception:
        planet_xy = {}
    for _sid, body in _split(_brace(data, "ships")).items():
        fl = re.search(r"\n\t\tfleet=(\d+)", body)
        ds = re.search(r"design=(\d+)", body)
        coord = re.search(r"coordinate=\s*\{([^}]+)\}", body)
        orig = re.search(r"origin=(\d+)", coord.group(1)) if coord else None
        if not (fl and ds and orig):
            continue
        sz = designs.get(ds.group(1))
        sys_id = orig.group(1)
        if sys_id == "4294967295":
            continue
        owner = fleet_owner.get(fl.group(1))
        ctype = ctypes.get(owner or "", "")
        if sz == "mining_station" and ctype == "drone":
            drone_mines[sys_id] += 1
            systems.add(sys_id)
            if planet_xy:
                xm = re.search(r"x=([-\d.]+)", coord.group(1))
                ym = re.search(r"y=([-\d.]+)", coord.group(1))
                if xm and ym:
                    pid = continuum_cp._nearest_planet(
                        float(xm.group(1)), float(ym.group(1)), sys_id, planet_xy, want="mine"
                    )
                    if pid:
                        drone_mine_planets.append(str(pid))
            continue
        spec = FAUNA_DESIGNS.get(sz)
        if not spec:
            continue
        design, kind, default_name = spec
        if ctype in GARRISON_CTYPES:
            kind = GARRISON_CTYPES[ctype]
        rec = fleets_out.setdefault(fl.group(1), {
            "sys": sys_id,
            "kind": kind,
            "garrison": ctype in GARRISON_CTYPES,
            "ships": defaultdict(int),
            "name": fleet_names.get(fl.group(1)) or default_name,
        })
        rec["sys"] = sys_id
        rec["kind"] = kind
        rec["ships"][design] += 1
        systems.add(sys_id)
    return {
        "systems": systems,
        "fleets": list(fleets_out.values()),
        "drone_mines": dict(drone_mines),
        "drone_mine_planets": drone_mine_planets,
    }


def _create_countries():
    return r'''			create_crystal_country = yes
			create_drone_country = yes
			create_amoeba_country = yes
			create_tiyanki_country = yes
			create_cloud_country = yes
			if = {
				limit = { NOT = { exists = event_target:voidworms_country } }
				create_country = {
					name = "NAME_Voidworms"
					type = voidworms
					flag = {
						icon = { category = "zoological" file = "flag_zoological_1.dds" }
						background = { category = "backgrounds" file = "00_solid.dds" }
						colors = { "black" "black" "null" "null" }
					}
					effect = { save_global_event_target_as = voidworms_country }
				}
			}
			create_tiyanki_garrison_country = yes
			if = {
				limit = { NOT = { exists = event_target:amoeba_garrison_country } }
				create_country = {
					name = "NAME_Space_Amoebas_United"
					type = amoeba_garrison
					flag = {
						icon = { category = "zoological" file = "flag_zoological_1.dds" }
						background = { category = "backgrounds" file = "00_solid.dds" }
						colors = { "black" "black" "null" "null" }
					}
					effect = { save_global_event_target_as = amoeba_garrison_country }
				}
			}
'''


def _owner_target(kind, garrison):
    if kind == "crystal":
        return "crystal_country"
    if kind == "drone":
        return "drone_country"
    if kind == "gardener":
        return "gardener_country"
    if kind == "cloud":
        return "cloud_country"
    if kind == "voidworms":
        return "voidworms_country"
    if kind == "coven":
        return "shroudwalker_enclave_country"
    if kind == "amoeba":
        return "amoeba_garrison_country" if garrison else "amoeba_country"
    if kind == "tiyanki":
        return "tiyanki_garrison_country" if garrison else "tiyanki_country"
    if kind == "lost_swarm":
        return "lost_swarm_country"
    if kind == "hillos":
        return "hillos_country"
    return None


def _ensure_special_countries(kinds):
    bits = []
    if "gardener" in kinds:
        bits.append(r'''			if = {
				limit = { NOT = { exists = event_target:gardener_country } }
				create_species = {
					name = random
					class = MACHINE
					portrait = random
					traits = { ideal_planet_class = pc_gaia trait = trait_mechanical }
				}
				last_created_species = { save_global_event_target_as = wenkwort_audon }
				create_country = {
					name = "NAME_Gardeners"
					type = drone
					species = event_target:wenkwort_audon
					flag = {
						icon = { category = "ornate" file = "flag_ornate_18.dds" }
						background = { category = "backgrounds" file = "v.dds" }
						colors = { "red_orange" "green" "null" "null" }
					}
					effect = { save_global_event_target_as = gardener_country }
				}
			}
''')
    if "lost_swarm" in kinds:
        bits.append(r'''			if = {
				limit = { NOT = { exists = event_target:lost_swarm_country } }
				create_country = {
					name = "NAME_Lost_Swarm"
					type = faction
					flag = {
						icon = { category = "domination" file = "domination_14.dds" }
						background = { category = "backgrounds" file = "00_solid.dds" }
						colors = { "black" "black" "null" "null" }
					}
					effect = {
						set_graphical_culture = swarm_01
						set_country_flag = lost_swarm
						save_global_event_target_as = lost_swarm_country
					}
				}
			}
''')
    if "hillos" in kinds:
        bits.append(r'''			if = {
				limit = { NOT = { exists = event_target:hillos_country } }
				create_country = {
					name = "NAME_Hillos_fleet"
					type = neutral_faction
					flag = {
						icon = { category = "special" file = "unknown.dds" }
						background = { category = "backgrounds" file = "00_solid.dds" }
						colors = { "black" "black" "null" "null" }
					}
					effect = { save_global_event_target_as = hillos_country }
				}
			}
''')
    return "".join(bits)


def _fleet_name_token(name):
    if not name:
        return ""
    s = str(name)
    if re.match(r"^[A-Za-z][A-Za-z0-9_]*$", s) and "_" in s:
        return f"name = {s}"
    return f'name = "{s.replace(chr(34), "")}"'


def _fleet_block(kind, garrison, ships, flag, fleet_name=None):
    owner = _owner_target(kind, garrison)
    if not owner or not ships:
        return ""
    nest_n = ships.pop("NAME_Voidworms_Starbase", 0) if kind == "voidworms" else 0
    nidus_n = ships.pop("NAME_Crystal_Nidus", 0) if kind == "crystal" else 0
    avatar_n = ships.pop("NAME_Shroud_Avatar", 0) if kind == "coven" else 0
    whiles = []
    for design, n in sorted(ships.items()):
        if n < 1:
            continue
        ship_nm = "\"\"" if kind not in ("gardener", "hillos") else f'"{design}"'
        whiles.append(
            f"								while = {{ count = {int(n)} create_ship = {{ name = {ship_nm} design = \"{design}\" }} }}"
        )
    garrison_set = "garrison = yes " if garrison or kind in ("voidworms",) else ""
    parts = []
    if whiles:
        inner = "\n".join(whiles)
        stance = "passive" if kind in ("tiyanki", "gardener", "coven", "hillos") else "aggressive"
        nline = _fleet_name_token(fleet_name)
        name_bit = f"\n								{nline}" if nline else ""
        parts.append(f"""			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				random_system_planet = {{
					limit = {{ is_star = no }}
					if = {{
						limit = {{ exists = event_target:{owner} }}
						event_target:{owner} = {{
							create_fleet = {{{name_bit}
								settings = {{ spawn_debris = no {garrison_set}}}
								effect = {{
									set_owner = prev
{inner}
									set_location = PREVPREV
									set_fleet_stance = {stance}
								}}
							}}
						}}
					}}
				}}
			}}""")
    for _ in range(int(nest_n)):
        parts.append(f"""			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				random_system_planet = {{
					limit = {{ is_star = no }}
					if = {{
						limit = {{ exists = event_target:voidworms_country }}
						event_target:voidworms_country = {{
							create_fleet = {{
								name = "NAME_Voidworms_Starbase"
								settings = {{ spawn_debris = no garrison = yes }}
								effect = {{
									set_owner = prev
									create_ship = {{ name = "NAME_Voidworms_Starbase" design = "NAME_Voidworms_Starbase" upgradable = no }}
									set_location = {{ target = PREVPREV distance = 40 angle = random }}
								}}
							}}
						}}
					}}
				}}
			}}""")
    if nidus_n:
        parts.append(f"""			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				random_system_planet = {{
					limit = {{ is_star = no }}
					if = {{
						limit = {{ exists = event_target:crystal_country }}
						event_target:crystal_country = {{
							create_fleet = {{
								name = "NAME_Crystal_Nidus"
								settings = {{ spawn_debris = no garrison = yes }}
								effect = {{
									set_owner = prev
									create_ship = {{ name = random design = "NAME_Crystal_Nidus" }}
									set_location = PREVPREV
								}}
							}}
						}}
					}}
				}}
			}}""")
    if avatar_n:
        parts.append(f"""			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				random_system_planet = {{
					limit = {{ is_star = yes }}
					if = {{
						limit = {{ exists = event_target:shroudwalker_enclave_country }}
						event_target:shroudwalker_enclave_country = {{
							create_fleet = {{
								name = "NAME_Psionic_Avatar"
								settings = {{ spawn_debris = no is_boss = yes can_upgrade = no }}
								effect = {{
									set_owner = prev
									set_fleet_flag = avatar_fleet
									create_ship = {{ name = "NAME_Avatar" design = "NAME_Shroud_Avatar" prefix = no upgradable = no }}
									set_location = PREVPREV
								}}
							}}
						}}
					}}
				}}
			}}""")
    return "\n".join(parts)


def write_fauna_events_file(output_path, snapshot):
    if not isinstance(snapshot, dict) or "fleets" not in snapshot:
        snapshot = {"systems": set(), "fleets": [], "drone_mines": {}}
    kinds = set()
    blocks = []
    for fl in snapshot.get("fleets") or []:
        kind = fl.get("kind")
        kinds.add(kind)
        flag = f"continuum_s{fl.get('sys')}"
        ships = dict(fl.get("ships") or {})
        txt = _fleet_block(kind, bool(fl.get("garrison")), ships, flag, fl.get("name"))
        if txt:
            blocks.append(txt)
    drone_planets = snapshot.get("drone_mine_planets") or []
    if drone_planets:
        for i, _pid in enumerate(drone_planets):
            blocks.append(f"""			every_galaxy_planet = {{
				limit = {{
					has_planet_flag = continuum_drone_mine_{i}
					NOT = {{ exists = mining_station }}
				}}
				if = {{
					limit = {{ exists = event_target:drone_country }}
					create_mining_station = {{ owner = event_target:drone_country }}
				}}
			}}""")
    else:
        drone_mine_counts = {}
        for sys_id, n_mines in (snapshot.get("drone_mines") or {}).items():
            if int(n_mines) > 0:
                drone_mine_counts[f"continuum_s{sys_id}"] = int(n_mines)
        for flag, n_mines in drone_mine_counts.items():
            blocks.append(f"""			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				while = {{
					count = {n_mines}
					random_system_planet = {{
						limit = {{
							has_deposit_for = shipclass_mining_station
							NOT = {{ exists = mining_station }}
						}}
						if = {{
							limit = {{ exists = event_target:drone_country }}
							create_mining_station = {{ owner = event_target:drone_country }}
						}}
					}}
				}}
			}}""")
    inner = _create_countries() + _ensure_special_countries(kinds) + ("\n".join(blocks) + "\n" if blocks else "			# no fauna snapshot\n")
    content = f"""namespace = continuum_npc
event = {{
	id = continuum_npc.2
	is_triggered_only = yes
	hide_window = yes
	immediate = {{
		if = {{
			limit = {{ has_global_flag = continuum_fauna_done }}
		}}
		else = {{
			set_global_flag = continuum_fauna_done
{inner}		}}
	}}
}}
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    n_sys = len(snapshot.get("systems") or [])
    n_fleets = len(snapshot.get("fleets") or [])
    n_ships = sum(sum((fl.get("ships") or {}).values()) for fl in snapshot.get("fleets") or [])
    n_ships += len(snapshot.get("drone_mine_planets") or []) or sum(
        int(n) for n in (snapshot.get("drone_mines") or {}).values()
    )
    print(f"Fauna snapshot: {n_fleets} fleets in {n_sys} systems, {n_ships} Pre ships/stations queued")
    return list(snapshot.keys())
