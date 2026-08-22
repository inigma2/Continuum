"""0.7: parse Pre default empires / primitives, pick a new-polity start, nudge borders, emit spawn."""
import random
import re
import zipfile

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


def _name_from_block(body):
    m = re.search(r'\n\t\tname=\s*\n\t\t\{\s*\n\t\t\tkey="([^"]+)"(.*?)\}', body, re.DOTALL)
    if not m:
        m = re.search(r'key="([^"]+)"', body[:1200])
        return m.group(1) if m else "Unknown"
    key, rest = m.group(1), m.group(2)
    if key.startswith("NAME_") or key.startswith("SPEC_"):
        return key
    if "literal=yes" in rest[:80] or key not in ("%ADJECTIVE%", "%ADJ%", "AofB"):
        return key.replace("_", " ")
    vm = re.search(r'value=\s*\{\s*\n\s*key="([^"]+)"', rest)
    if vm:
        v = vm.group(1)
        if v.startswith("NAME_") or v.startswith("SPEC_"):
            return v
        return v.replace("SPEC_", "").replace("_", " ")
    return key.replace("_", " ")


def script_name(name):
    if name.startswith("NAME_") or name.startswith("SPEC_"):
        return name
    return f'"{name.replace(chr(34), "")}"'


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
            spawn[sid] = {"base": 2, "modifiers": [{"add": 25, "class": cl}]}
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


def write_intro_and_empire_events(events_dir, loc_entries, plan, empires, species, owned, capitals, trait_keys, civic_keys, ethic_keys):
    import os
    intro_path = os.path.join(events_dir, "continuum_intro_events.txt")
    crisis_flag = ""
    if plan.get("had_crisis"):
        crisis_flag = "\t\t\tset_global_flag = continuum_had_crisis\n"
    with open(intro_path, "w", encoding="utf-8") as f:
        f.write(f"""namespace = continuum_intro
country_event = {{
	id = continuum_intro.1
	is_triggered_only = yes
	title = continuum_intro_title
	desc = {{
		trigger = {{ has_global_flag = continuum_had_crisis }}
		text = continuum_intro_crisis
	}}
	desc = {{
		trigger = {{
			NOT = {{ has_global_flag = continuum_had_crisis }}
			solar_system = {{ has_star_flag = continuum_was_primitive }}
		}}
		text = continuum_intro_primitive
	}}
	desc = {{
		trigger = {{
			NOT = {{ has_global_flag = continuum_had_crisis }}
			solar_system = {{ has_star_flag = continuum_border }}
		}}
		text = continuum_intro_border
	}}
	desc = {{
		trigger = {{
			NOT = {{ has_global_flag = continuum_had_crisis }}
			solar_system = {{ has_star_flag = continuum_unowned }}
		}}
		text = continuum_intro_new
	}}
	desc = continuum_intro_present
	picture = GFX_evt_throne_room
	show_sound = event_default
	trigger = {{ is_ai = no }}
	immediate = {{
		if = {{
			limit = {{ has_global_flag = continuum_intro_done }}
		}}
		else = {{
			set_global_flag = continuum_intro_done
{crisis_flag}		}}
	}}
	option = {{
		name = continuum_intro_ok
	}}
}}
""")

    emp_path = os.path.join(events_dir, "continuum_empire_events.txt")
    chunks = []
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
					set_country_flag = continuum_pre_empire
					give_technology = {{ tech = tech_starbase_1 message = no }}
					give_technology = {{ tech = tech_starbase_2 message = no }}
					give_technology = {{ tech = tech_corvettes message = no }}
					refresh_auto_generated_ship_designs = yes
					add_resource = {{ energy = 1000 minerals = 1000 food = 1000 alloys = 500 influence = 200 }}
				}}
			}}
			every_system = {{
				limit = {{ has_star_flag = {flag} }}
				every_system_planet = {{
					limit = {{
						OR = {{
							has_planet_flag = continuum_emp_{idx}_homeworld
							has_planet_flag = continuum_emp_{idx}_colony
						}}
					}}
					if = {{
						limit = {{ exists = event_target:continuum_emp_{idx} }}
						create_colony = {{
							owner = event_target:continuum_emp_{idx}
							species = event_target:continuum_sp_{idx}
						}}
						generate_start_pops = yes
					}}
				}}
			}}
			every_system = {{
				limit = {{
					has_star_flag = {flag}
					NOT = {{ exists = space_owner }}
					NOT = {{ exists = starbase }}
				}}
				if = {{
					limit = {{ exists = event_target:continuum_emp_{idx} }}
					create_starbase = {{
						size = starbase_starport
						owner = event_target:continuum_emp_{idx}
					}}
				}}
			}}
			if = {{
				limit = {{ exists = event_target:continuum_emp_{idx} }}
				random_system = {{
					limit = {{ has_star_flag = {cap_flag} }}
					event_target:continuum_emp_{idx} = {{
						create_fleet = {{
							name = continuum_home_fleet
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								while = {{
									count = 3
									create_ship = {{ name = random random_existing_design = corvette }}
								}}
								set_location = prevprev
							}}
						}}
					}}
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
        sp_name = script_name(sp.get("name") or "Unknown")
        ethics = [e for e in (emp.get("ethics") or []) if e.startswith("ethic_")]
        if ethic_keys:
            ethics = [e for e in ethics if e in ethic_keys]
        if not ethics:
            ethics = ["ethic_fanatic_materialist"]
        fe_n = emp.get("fallen_n") or FALLEN_BY_ETHIC.get(ethics[0], "1")
        gfx = emp.get("graphical_culture") or f"fallen_empire_0{fe_n}"
        designs = FALLEN_DESIGNS.get(fe_n) or FALLEN_DESIGNS["1"]
        design_lines = "\n".join(f"					add_global_ship_design = \"{d}\"" for d in designs)
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys][:8]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits) or "\t\t\t\ttrait = trait_adaptive"
        colors = emp.get("flag_colors") or ["black", "black", "black", "null"]
        while len(colors) < 4:
            colors.append("null")
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
				name = {name}
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
					add_resource = {{ minerals = 10000 energy = 10000 food = 1000 influence = 500 }}
{design_lines}
				}}
			}}
			every_system = {{
				limit = {{ has_star_flag = continuum_fe_{idx} }}
				every_system_planet = {{
					limit = {{
						OR = {{
							has_planet_flag = continuum_fe_{idx}_homeworld
							has_planet_flag = continuum_fe_{idx}_colony
						}}
					}}
					if = {{
						limit = {{ exists = event_target:continuum_fe_{idx} }}
						create_colony = {{
							owner = event_target:continuum_fe_{idx}
							species = event_target:continuum_fe_sp_{idx}
						}}
						generate_start_pops = yes
					}}
				}}
			}}
			every_system = {{
				limit = {{
					has_star_flag = continuum_fe_{idx}
					NOT = {{ exists = space_owner }}
					NOT = {{ exists = starbase }}
				}}
				if = {{
					limit = {{ exists = event_target:continuum_fe_{idx} }}
					create_starbase = {{
						size = starbase_citadel
						owner = event_target:continuum_fe_{idx}
					}}
				}}
			}}
""")

    for idx, mar in enumerate(plan.get("marauders") or []):
        if not mar.get("home_system"):
            continue
        sp = species.get(str(mar.get("founder_species"))) or {}
        name = script_name(mar.get("name") or "Unknown")
        sp_name = script_name(sp.get("name") or "Unknown")
        n = mar.get("marauder_n") or str(idx + 1)
        traits = [t for t in (sp.get("traits") or []) if not trait_keys or t in trait_keys][:8]
        trait_lines = "\n".join(f"\t\t\t\ttrait = {t}" for t in traits) or "\t\t\t\ttrait = trait_rapid_breeders"
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
					name = {name}
					type = dormant_marauders
					civics = {{ civic = civic_anarcho_tribalism }}
					origin = origin_default
					species = last_created_species
					ethos = {{ ethic = ethic_fanatic_militarist ethic = ethic_xenophobe }}
					ignore_initial_colony_error = yes
					day_zero_contact = no
					exclude_day_zero_contact = event_target:continuum_human
					effect = {{
						save_global_event_target_as = continuum_mar_{idx}
						set_graphical_culture = pirate_01
						set_country_flag = marauder_{n}
						set_country_flag = continuum_pre_marauder
						create_ship_design = {{ design = "NAME_Marauder_Starbase" }}
						add_ship_design = last_created_design
					}}
				}}
				last_created_country = {{
					create_fleet = {{
						settings = {{ spawn_debris = no }}
						effect = {{
							set_owner = prev
							create_ship = {{ name = random design = "NAME_Warrior_Freehold" graphical_culture = pirate_01 }}
							set_location = {{ target = prevprev distance = 80 }}
						}}
					}}
				}}
			}}
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
        if not str(origin).startswith("origin_"):
            origin = "origin_default_pre_ftl"
        colors = prim.get("flag_colors") or ["turquoise", "green", "null", "null"]
        while len(colors) < 4:
            colors.append("null")
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
					}}
				}}
				create_colony = {{
					owner = last_created_country
					species = last_created_species
				}}
				}}
			}}
""")

    chunks.append("		}\n	}\n}\n")
    with open(emp_path, "w", encoding="utf-8") as f:
        f.write("".join(chunks))

    loc_entries["continuum_home_fleet"] = "Home Fleet"
    loc_entries["continuum_intro_title"] = "A New Polity"
    loc_entries["continuum_intro_ok"] = "Begin"
    loc_entries["continuum_intro_present"] = "A thousand years have passed. The old empires still hold the cores of their space. You are a new political entity in their galaxy — not their heir."
    loc_entries["continuum_intro_new"] = "A thousand years have passed. Your people have only now reached FTL. The powers of the last age still occupy their cores. The rim, and the gaps between them, are yours to claim."
    loc_entries["continuum_intro_border"] = "A thousand years have passed. You rise on a frontier the old empires never fully settled. They still exist, inland. You are a new state on their doorstep."
    loc_entries["continuum_intro_primitive"] = "A thousand years have passed. Your world was pre-FTL in the last age. You have reached the stars. The empires that once mapped this sky still exist. You do not serve them."
    loc_entries["continuum_intro_crisis"] = "A thousand years have passed since the crisis that scarred this galaxy. The old empires that survived still hold their cores. You are a new polity in the wreckage — survivors, or something that grew in the dark."
    loc_entries["opinion_continuum_ceasefire"] = "Lingering Hostility"


def _tag_owned_planets(countries, colony_to_planet, prefix, flags, planet_flags, owned, capitals):
    for idx, emp in enumerate(countries):
        for sid in owned.get(emp["id"], []):
            flags.setdefault(sid, []).append(f"{prefix}_{idx}")
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


def build_empire_plan(save_path, galaxy_data, stars, planets, argv_start=None):
    countries, species, had_crisis, _data, colony_to_planet = load_empire_save_data(save_path)
    p2s = planet_to_system_map(stars)
    attach_primitive_systems(countries, p2s, colony_to_planet)
    habitables = system_habitables(galaxy_data, planets)
    blocked = blocked_systems(galaxy_data)
    fe_owned, fe_caps = ownership_from_planets(countries, p2s, colony_to_planet, types=("fallen_empire",))
    for syss in fe_owned.values():
        blocked.update(syss)
    marauders = [c for c in countries if c["type"] == "dormant_marauders"]
    marauder_homes = {}
    for sys in galaxy_data:
        sid = str(sys.get("id"))
        for fl in sys.get("flags") or []:
            sfl = str(fl)
            if sfl.startswith("marauder_capital_"):
                marauder_homes[sfl] = sid
                blocked.add(sid)
            if sfl in ("marauder_system", "marauder_capital_1", "marauder_capital_2", "marauder_capital_3"):
                blocked.add(sid)
    owned, capitals = ownership_from_planets(countries, p2s, colony_to_planet, types=("default",))
    for cid, syss in list(owned.items()):
        cap = capitals.get(cid)
        owned[cid] = [s for s in syss if s not in blocked or s == cap]
    lanes = hyperlane_map(galaxy_data)
    rng = random.Random(int(re.sub(r"\D", "", save_path)[-8:] or "1") or 1)
    owned = nudge_borders(owned, capitals, lanes, blocked, set(), rng)
    tags, spawn_weights = classify_systems(owned, capitals, lanes, blocked, habitables, countries, species)
    defaults = [c for c in countries if c["type"] == "default" and owned.get(c["id"])]
    fallens = [c for c in countries if c["type"] == "fallen_empire" and fe_owned.get(c["id"])]
    primitives = [c for c in countries if c["type"] == "primitive" and c.get("system_id")]
    flags = {}
    planet_flags = {}
    for sid, flist in tags.items():
        flags.setdefault(sid, []).extend(flist)
    _tag_owned_planets(defaults, colony_to_planet, "continuum_emp", flags, planet_flags, owned, capitals)
    _tag_owned_planets(fallens, colony_to_planet, "continuum_fe", flags, planet_flags, fe_owned, fe_caps)
    for idx, prim in enumerate(primitives):
        sid = prim.get("system_id")
        if sid:
            flags.setdefault(sid, []).append(f"continuum_prim_{idx}")
        cap_col = prim.get("capital") or ((prim.get("owned_planets") or [None])[0])
        cap_pid = colony_to_planet.get(str(cap_col)) if cap_col is not None else None
        if cap_pid:
            planet_flags.setdefault(str(cap_pid), []).append(f"continuum_prim_{idx}_homeworld")
    for idx, mar in enumerate(marauders):
        n = mar.get("marauder_n") or str(idx + 1)
        cap_sid = marauder_homes.get(f"marauder_capital_{n}") or marauder_homes.get(f"marauder_capital_{idx + 1}")
        if cap_sid:
            flags.setdefault(cap_sid, []).append(f"continuum_mar_{idx}")
            flags.setdefault(cap_sid, []).append(f"continuum_mar_{idx}_capital")
            mar["home_system"] = cap_sid
            mar["marauder_n"] = n
    enclave_names = {}
    for c in countries:
        if c["type"] != "enclave":
            continue
        civics = set(c.get("civics") or [])
        if "civic_shroudwalker_enclave" in civics:
            enclave_names["shroudwalker"] = c.get("name")
        elif "civic_salvager_enclave" in civics:
            enclave_names["salvager"] = c.get("name")
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
