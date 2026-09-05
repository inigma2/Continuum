"""Continuum Present slices B–F: parse extra save fields and emit numbered events."""
import re

MILITARY_SIZES = frozenset({
    "corvette", "frigate", "destroyer", "cruiser", "battleship", "titan",
    "small_ship_fallen_empire", "large_ship_fallen_empire", "massive_ship_fallen_empire",
})
CIVILIAN_SIZES = frozenset({"science", "constructor", "colonizer", "transport"})
STATION_SIZES = {
    "mining_station": "mining",
    "research_station": "research",
    "observation_station": "observe",
}
STAR_CLASS_MARKERS = ("_star", "black_hole", "pulsar", "neutron", "quasar")
HABITABLE_CLASS_MARKERS = (
    "continental", "ocean", "tropical", "arid", "desert", "tundra", "arctic",
    "alpine", "savannah", "gaia", "habitat", "ringworld_habitable", "city",
    "ecumenopolis",
)
GESTALT_AUTHS = frozenset({"auth_hive_mind", "auth_machine_intelligence"})
SKIP_MEGA_TYPES = frozenset({
    "lgate_base", "gateway_ruined", "dyson_sphere_ruined",
    "spy_orb_ruined", "interstellar_assembly_ruined", "orbital_ring_ruined",
    "ring_world_ruined", "think_tank_ruined", "mega_art_installation_ruined",
    "spy_orb_ruined_empty",
})
FORMAT_NAME_KEYS = frozenset({
    "PREFIX_NAME_FORMAT", "SUFFIX_NAME_FORMAT",
    "STARBASE_STATION_NAME_FORMAT", "STARBASE_STATION_NAME_FORMAT_NON_PRIMARY",
    "shipclass_starbase_name", "shipclass_mining_station_name",
    "shipclass_research_station_name", "shipclass_science_ship_name",
    "FLEET_MANAGER_FLEET_NAME_FORMAT", "ASTEROID_NAME_FORMAT",
    "PLANET_NAME_FORMAT", "SUBPLANET_NAME_FORMAT",
    "REGNAL_NAME_FORMAT_NUMBERED", "%SEQ%",
})
LEADER_CLASSES = frozenset({"official", "scientist", "commander"})
# Soldier/warrior jobs spawn these on colonies. create_army of the same type
# stacks on top and doubles garrison (Weeer 14 -> 28, UI ~623 -> ~1194).
JOB_DEFENSE_ARMY_TYPES = frozenset({
    "defense_army",
    "machine_defense",
    "robotic_defense_army",
    "undead_defense_army",
    "psionic_defense_army",
})

_LOC = {}
FORMAT_TOP = frozenset({
    "AofB", "AofBpfx", "AassocB", "%ADJECTIVE%", "%ADJ%",
    "SEQ", "SEQ2", "SEQ3", "%SEQ%",
    "%LEADER_1%", "%LEADER_2%", "%LEADER_3%",
})
FORMAT_SKIP_KEYS = frozenset({
    "adjective", "1", "2", "3", "%ADJECTIVE%", "%ADJ%",
    "AofB", "AofBpfx", "AassocB", "SEQ", "SEQ2", "SEQ3", "%SEQ%",
    "key", "value", "variables", "THIS", "This",
    "prefix", "suffix", "NAME", "PARENT", "NUMERAL", "fmt", "format", "num",
})


def set_loc_data(loc_data):
    global _LOC
    _LOC = loc_data or {}


def loc_lookup(key):
    if not key:
        return None
    val = _LOC.get(key)
    if not val or "$" in val or val.startswith("%"):
        return None
    return val


def display_key(key):
    """Resolve a save name key to a player-visible string via vanilla loc."""
    if not key:
        return ""
    looked = loc_lookup(key)
    if looked:
        return looked
    s = str(key)
    if "_CHR_" in s:
        s = s.split("_CHR_", 1)[-1]
    s = re.sub(
        r"^(NAME_|SPEC_|PRESCRIPTED_species_name_|PRESCRIPTED_ruler_name_|"
        r"PRESCRIPTED_adjective_|PRESCRIPTED_ship_prefix_|EMPIRE_DESIGN_|PRESCRIPTED_)",
        "",
        s,
    )
    s = re.sub(r"_(planet|system|star|moon)$", "", s, flags=re.I)
    return s.replace("_", " ").strip() or str(key)


def apply_adjective(noun, rest):
    """English adj_NN* inflection: Faller + Shard -> Falleran Shard (adj_NNr *ran)."""
    s = str(noun or "").strip()
    rest = str(rest or "").strip()
    if not s:
        return rest
    # Apostrophe names are already a display form (Rihi'Nar, not Rihi'Naran).
    if "'" in s or "’" in s:
        return f"{s} {rest}".strip()
    rules = []
    for k, v in (_LOC or {}).items():
        if not str(k).startswith("adj_NN") or not v:
            continue
        suffix = str(k)[6:]
        rules.append((len(suffix), suffix, v))
    rules.sort(reverse=True)
    tmpl = None
    stem = s
    for _n, suffix, v in rules:
        if suffix and s.lower().endswith(suffix.lower()):
            stem = s[: len(s) - len(suffix)]
            tmpl = v
            break
    if not tmpl:
        return f"{s} {rest}".strip()
    out = tmpl.replace("*", stem).replace("$1$", rest)
    return re.sub(r"\s+", " ", out).strip()


def name_from_keys(keys):
    """Compose a display name from a Clausewitz name-block key list."""
    if not keys:
        return "Unknown"
    top = keys[0]
    template = _LOC.get(top, "") if top else ""
    is_format = top in FORMAT_TOP or ("$" in template)
    if is_format:
        parts = []
        for k in keys[1:]:
            if k in FORMAT_SKIP_KEYS:
                continue
            p = display_key(k)
            if p and p not in parts:
                parts.append(p)
        if top in ("%ADJECTIVE%", "%ADJ%") and parts:
            return apply_adjective(parts[0], " ".join(parts[1:]))
        if "$1$" in template:
            out = template
            for i, p in enumerate(parts, 1):
                out = out.replace(f"${i}$", p)
            out = re.sub(r"\$\d+\$", "", out).strip()
            if out:
                return out
        if top == "AofB" and len(parts) >= 2:
            return f"{parts[0]} of {parts[1]}"
        return " ".join(parts) if parts else display_key(top)
    return display_key(top)


def _brace_section(data, key):
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


def _split_entries(section, tab="\t"):
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


def _script_token(name):
    if not name:
        return '"Unknown"'
    if re.match(r"^[A-Za-z][A-Za-z0-9_]*$", name) and "_" in name:
        return name
    return '"' + name.replace('"', "") + '"'


def _is_format_name(name):
    """True if this is a loc format template, not a display name."""
    if not name:
        return True
    if name in FORMAT_NAME_KEYS or name.startswith("%"):
        return True
    if name.startswith("shipclass_"):
        return True
    if name.endswith("_FORMAT") or name.endswith("_NAME_FORMAT"):
        return True
    return False


def pretty_ship_name(nm):
    """Player-visible ship name. Keep loc text as-is (MACHINE4 is literally composite717)."""
    if not nm or _is_format_name(nm):
        return None
    s = str(nm)
    looked = loc_lookup(s)
    if looked and "$" not in looked:
        return looked
    if "SHIP_" in s:
        return s.split("SHIP_")[-1].replace("_", " ")
    return s


def _is_loc_key(s):
    return bool(s) and bool(re.match(r"^[A-Za-z][A-Za-z0-9_]*$", str(s)))


def ship_name_token(nm):
    """Unquoted loc keys so PREFIX_NAME_FORMAT can apply. Quoted display strings otherwise."""
    if not nm or _is_format_name(nm):
        return "random"
    s = str(nm)
    if _is_loc_key(s):
        return s
    pretty = pretty_ship_name(s)
    if not pretty:
        return "random"
    if _is_loc_key(pretty):
        return pretty
    return '"' + pretty.replace('"', "") + '"'


def _ordinal(n):
    n = int(n)
    if 10 <= n % 100 <= 20:
        suf = "th"
    else:
        suf = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"


def _seq_fleet_name(keys):
    """%SEQ% + fmt=TOX4_fleet_names + num=2 -> '2nd Voidlurcher'."""
    if not keys:
        return None
    fmt = None
    num = None
    i = 0
    while i < len(keys):
        k = keys[i]
        if k in ("fmt", "format") and i + 1 < len(keys):
            fmt = keys[i + 1]
            i += 2
            continue
        if k == "num" and i + 1 < len(keys):
            num = keys[i + 1]
            i += 2
            continue
        i += 1
    if not fmt:
        for k in keys:
            if "fleet_names" in k:
                fmt = k
                break
    if not fmt:
        return None
    template = _LOC.get(fmt) or ""
    n = None
    if num and str(num).isdigit():
        n = int(num)
    else:
        for k in keys:
            if k.isdigit():
                n = int(k)
                break
    if n is None:
        n = 1
    out = (
        template.replace("$ORD$", _ordinal(n))
        .replace("$O$", _ordinal(n))
        .replace("$NUM$", str(n))
        .replace("$FLEET$", _ordinal(n))
    )
    out = re.sub(r"£[^£]+£", "", out).strip()
    if out and "$" not in out:
        return out
    return None


def _fleet_name_line(nm):
    if not nm or _is_format_name(nm):
        return ""
    s = str(nm)
    if _is_loc_key(s):
        return f"name = {s}"
    return f'name = "{s.replace(chr(34), "")}"'


_PREFIX_SKIP = frozenset({"of", "the", "and", "a", "an", "de", "da", "von", "van"})


def stellaris_acronym(pretty):
    """Vanilla %ACRONYM%: first letter of each word, pad to 3 with last letter of last word.

    Loc documents AofBpfx 'Empire Sol' -> 'ESL'. So 'Aramathi Mandate' -> 'AME'.
    """
    words = [w for w in re.findall(r"[A-Za-z0-9]+", pretty or "") if w.lower() not in _PREFIX_SKIP]
    if not words:
        return ""
    if len(words) == 1:
        w = words[0]
        token = (w[:3] if len(w) >= 3 else w).upper()
        return token if re.match(r"^[A-Za-z][A-Za-z0-9]{0,7}$", token) else ""
    letters = [w[0].upper() for w in words]
    if len(letters) < 3:
        last = words[-1]
        if last:
            letters.append(last[-1].upper())
    token = "".join(letters[:3])
    return token if token and re.match(r"^[A-Za-z][A-Za-z0-9]{0,7}$", token) else ""


def resolved_ship_prefix(emp):
    """Literal prefix, PRESCRIPTED loc, or vanilla %ACRONYM% from the prefix-format base."""
    raw = str(emp.get("ship_prefix") or "").replace('"', "").strip()
    if raw.startswith("PRESCRIPTED_"):
        looked = loc_lookup(raw)
        if looked and not looked.startswith("%"):
            raw = looked.strip()
        else:
            raw = ""
    if raw and not raw.startswith("%"):
        token = re.sub(r"[^A-Za-z0-9]", "", raw)
        return token[:8] if token else ""
    base = str(emp.get("acronym_base") or "").strip()
    if not base:
        name = str(emp.get("name") or "")
        if name.startswith(("NAME_", "SPEC_", "EMPIRE_DESIGN_", "PRESCRIPTED_", "%")) or " " not in name:
            base = loc_lookup(name) or display_key(name) or name
        else:
            base = name
    return stellaris_acronym(base)


def _nested_section(blob, key):
    if not blob:
        return ""
    m = re.search(rf"\n[\t ]*{re.escape(key)}\s*=\s*\n?[\t ]*\{{", blob)
    if not m:
        return ""
    start = blob.find("{", m.start())
    if start < 0:
        return ""
    depth = 0
    for j in range(start, len(blob)):
        if blob[j] == "{":
            depth += 1
        elif blob[j] == "}":
            depth -= 1
            if depth == 0:
                return blob[start + 1:j]
    return ""


def _is_star_class(pc):
    if not pc:
        return True
    return any(s in pc for s in STAR_CLASS_MARKERS)


def _is_clearly_habitable(pc):
    if not pc:
        return False
    return any(s in pc for s in HABITABLE_CLASS_MARKERS)


def _is_gestalt(emp):
    auth = emp.get("authority") or ""
    if auth in GESTALT_AUTHS:
        return True
    if "ethic_gestalt_consciousness" in (emp.get("ethics") or []):
        return True
    return any("hive" in c or "machine" in c for c in (emp.get("civics") or []))


def _valid_tradition(key):
    if not key:
        return False
    return bool(re.match(r"^(tradition|tr)_[A-Za-z0-9_]+$", key))


def _valid_site_type(key):
    return bool(key) and bool(re.match(r"^[A-Za-z][A-Za-z0-9_]*$", key))


def _army_type_key(typ):
    if typ and re.match(r"^[A-Za-z][A-Za-z0-9_]*$", typ):
        return typ
    return "assault_army"


def _is_job_defense_army(typ):
    return _army_type_key(typ) in JOB_DEFENSE_ARMY_TYPES


def _is_councilor_trait(trait):
    t = str(trait or "")
    if "councilor" in t or "ruler_" in t:
        return True
    stem = re.sub(r"_\d+$", "", t)
    return stem in {
        "leader_trait_principled",
        "leader_trait_frontier_spirit",
        "leader_trait_intellectual",
        "leader_trait_logistic_understanding",
        "leader_trait_eye_for_talent",
        "leader_trait_fertility_preacher",
        "leader_trait_architectural_sense",
        "leader_trait_industrialist",
        "leader_trait_charismatic",
        "leader_trait_warlike",
        "leader_trait_recruiter",
        "leader_trait_feedback_loop",
    }


def _ship_display_name(body):
    m = re.search(r'key="NAME"\s*value=\s*\{\s*key="([^"]+)"', body)
    if m and not _is_format_name(m.group(1)):
        return m.group(1)
    m = re.search(r'\n\t\tname=\s*\{\s*key="([^"]+)"', body)
    if m and not _is_format_name(m.group(1)):
        return m.group(1)
    return None


def _fleet_display_name(body):
    keys = re.findall(r'key="([^"]+)"', body[:1000])
    if "%SEQ%" in keys or "SEQ" in keys or any("fleet_names" in k for k in keys):
        seq = _seq_fleet_name(keys)
        if seq:
            return seq
    m = re.search(r'key="NAME"\s*value=\s*\{\s*key="([^"]+)"', body[:800])
    if m and not _is_format_name(m.group(1)):
        return m.group(1)
    m = re.search(r'name=\s*\{\s*key="([^"]+)"', body[:500])
    if m and not _is_format_name(m.group(1)):
        return m.group(1)
    if keys:
        composed = name_from_keys(keys)
        if composed and composed != "Unknown" and not _is_format_name(composed):
            return composed
    return None


def _planet_xy(data):
    chunk = _brace_section(data, "planets")
    inner = _brace_section("\nplanet=\n{" + chunk, "planet") if False else ""
    i = data.find("\nplanets=\n")
    if i < 0:
        return {}
    # planets={ planet={ 0={ } } }
    pstart = data.find("planet=", i)
    if pstart < 0:
        return {}
    brace = data.find("{", pstart)
    depth = 0
    end = brace
    for k in range(brace, len(data)):
        if data[k] == "{":
            depth += 1
        elif data[k] == "}":
            depth -= 1
            if depth == 0:
                end = k
                break
    blocks = _split_entries(data[brace + 1:end], tab="\t\t")
    out = {}
    for pid, body in blocks.items():
        c = re.search(r"coordinate=\s*\{([^}]*)\}", body)
        pc = re.search(r'planet_class="([^"]+)"', body)
        if not c:
            continue
        x = re.search(r"x=([-\d.]+)", c.group(1))
        y = re.search(r"y=([-\d.]+)", c.group(1))
        orig = re.search(r"origin=(\d+)", c.group(1))
        if not (x and y and orig):
            continue
        out[pid] = (float(x.group(1)), float(y.group(1)), orig.group(1), (pc.group(1) if pc else ""), frozenset())
    dep_kinds = _planet_dep_kinds(data)
    for pid, kinds in dep_kinds.items():
        if pid in out:
            x, y, sid, pc, _ = out[pid]
            out[pid] = (x, y, sid, pc, kinds)
    return out


def _dep_kind(typ):
    t = typ or ""
    if any(x in t for x in ("physics", "society", "engineering", "astral")):
        return "res"
    if any(
        x in t
        for x in (
            "mineral", "energy", "alloy", "food", "trade", "exotic", "rare",
            "volatile", "zro", "dark_matter", "living_metal", "nanite",
            "mote", "gas", "betharian", "crystal",
        )
    ):
        return "mine"
    return None


def _planet_dep_kinds(data):
    dep_types = {}
    for did, body in _split_entries(_brace_section(data, "deposit"), tab="\t").items():
        m = re.search(r'type="([^"]+)"', body)
        if m:
            dep_types[did] = m.group(1)
    if not dep_types:
        for did, body in _split_entries(_brace_section(data, "deposits"), tab="\t").items():
            m = re.search(r'type="([^"]+)"', body)
            if m:
                dep_types[did] = m.group(1)
    kinds = {}
    i = data.find("\nplanets=\n")
    if i < 0:
        return kinds
    pstart = data.find("planet=", i)
    brace = data.find("{", pstart) if pstart >= 0 else -1
    if brace < 0:
        return kinds
    depth = 0
    end = brace
    for k in range(brace, len(data)):
        if data[k] == "{":
            depth += 1
        elif data[k] == "}":
            depth -= 1
            if depth == 0:
                end = k
                break
    for pid, body in _split_entries(data[brace + 1:end], tab="\t\t").items():
        blob = re.search(r"deposits=\s*\{([^}]*)\}", body)
        got = set()
        if blob:
            for tok in blob.group(1).split():
                if tok.isdigit():
                    k = _dep_kind(dep_types.get(tok, ""))
                    if k:
                        got.add(k)
        if got:
            kinds[pid] = frozenset(got)
    return kinds


def _station_body(pc):
    """Mining/research hosts: uninhabitable rocks and stars (energy mines)."""
    if not pc:
        return False
    if _is_clearly_habitable(pc):
        return False
    return True


def _nearest_planet(px, py, sys_id, planet_xy, mode="station", want=None):
    scored = []
    best_hab, hd = None, 1e18
    for pid, row in planet_xy.items():
        x, y, sid, pc = row[0], row[1], row[2], row[3]
        kinds = row[4] if len(row) > 4 else frozenset()
        if str(sid) != str(sys_id):
            continue
        if mode == "observe":
            if _is_star_class(pc):
                continue
        elif not _station_body(pc):
            continue
        d = (x - px) ** 2 + (y - py) ** 2
        scored.append((d, pid, kinds))
        if mode == "observe" and _is_clearly_habitable(pc) and d < hd:
            hd, best_hab = d, pid
    if mode == "observe" and best_hab is not None:
        return best_hab
    scored.sort()
    if want:
        for _d, pid, kinds in scored:
            if want in kinds:
                return pid
    return scored[0][1] if scored else None


def _slot_map(blob):
    """Parse `0=shipyard 1=solar_panel_network`."""
    if not blob:
        return {}
    out = {}
    for m in re.finditer(r"(\d+)=([A-Za-z0-9_]+)", blob):
        out[int(m.group(1))] = m.group(2)
    return out


def attach_cp_layers(data, countries):
    """Fill stations, named fleets, starbase fittings, megas, leaders, layouts."""
    by_id = {c["id"]: c for c in countries}
    for c in countries:
        c["named_military"] = []
        c["civilian_fleets"] = []
        c["mine_planets"] = []
        c["research_planets"] = []
        c["obs_planets"] = []
        c.setdefault("sb_fit", {})
        c.setdefault("mega_owned", [])
        c.setdefault("leaders", [])
        c.setdefault("ship_prefix", None)
        c.setdefault("acronym_base", "")
        c.setdefault("colony_layout", {})
        c.setdefault("overlord", None)
        c.setdefault("armies", [])
        c.setdefault("traditions", [])
        c.setdefault("federation_peers", [])
        c.setdefault("sites", [])
    colonized = set()
    for c in countries:
        colonized.update(str(p) for p in (c.get("colony_pop") or {}))

    csec = _split_entries(_brace_section(data, "country"), tab="\t")
    fleet_owner = {}
    for cid, body in csec.items():
        if cid not in by_id:
            continue
        pre_blob = _nested_section(body, "ship_prefix")
        if pre_blob:
            pre = re.search(r'key="([^"]+)"', pre_blob)
            raw_pre = pre.group(1) if pre else ""
            if raw_pre.startswith("%ACRONYM%") or raw_pre == "%ACRONYM%":
                keys = re.findall(r'key="([^"]+)"', pre_blob)
                rest = [k for k in keys if k not in ("%ACRONYM%", "base")]
                if rest:
                    by_id[cid]["acronym_base"] = name_from_keys(rest)
            elif raw_pre.startswith("PRESCRIPTED_"):
                looked = loc_lookup(raw_pre)
                if looked:
                    by_id[cid]["ship_prefix"] = looked
            elif raw_pre and not raw_pre.startswith("%"):
                by_id[cid]["ship_prefix"] = raw_pre
        ov = re.search(r"\n\t\toverlord=(\d+)", body)
        if ov and ov.group(1) != "4294967295":
            by_id[cid]["overlord"] = ov.group(1)
        trad_blob = _nested_section(body, "traditions")
        trads = [t for t in re.findall(r'"((?:tradition|tr)_[A-Za-z0-9_]+)"', trad_blob) if _valid_tradition(t)]
        by_id[cid]["traditions"] = trads[:20]
        fed = re.search(r"\n\t\tfederation=(\d+)", body)
        if fed and fed.group(1) != "4294967295":
            by_id[cid]["_federation"] = fed.group(1)
        start = body.find("owned_fleets=")
        chunk = body[start:] if start >= 0 else body
        for m in re.finditer(r"fleet=(\d+)", chunk):
            fid = m.group(1)
            if fid not in fleet_owner:
                fleet_owner[fid] = cid
        fm = re.search(r"\nfleets=\s*\n\t\t\{([^}]+)\}", body)
        if fm:
            for fid in fm.group(1).split():
                if fid.isdigit() and fid not in fleet_owner:
                    fleet_owner[fid] = cid

    designs = {}
    for did, body in _split_entries(_brace_section(data, "ship_design"), tab="\t").items():
        m = re.search(r'ship_size="([^"]+)"', body)
        if m:
            designs[did] = m.group(1)

    ships = _split_entries(_brace_section(data, "ships"), tab="\t")
    fleets = _split_entries(_brace_section(data, "fleet"), tab="\t")
    planet_xy = _planet_xy(data)

    fleet_ships = {}
    for sid, body in ships.items():
        fl = re.search(r"\n\t\tfleet=(\d+)", body)
        ds = re.search(r"design=(\d+)", body)
        if not fl or not ds:
            continue
        sz = designs.get(ds.group(1))
        fleet_ships.setdefault(fl.group(1), []).append((sid, sz, body))

    for fid, rows in fleet_ships.items():
        owner = fleet_owner.get(fid)
        if owner not in by_id:
            continue
        emp = by_id[owner]
        sizes = [sz for _sid, sz, _b in rows if sz]
        if any(sz == "military_station_small" for sz in sizes):
            emp.setdefault("defense_platforms", {})
            for _sid, sz, body in rows:
                if sz != "military_station_small":
                    continue
                c = re.search(r"origin=(\d+)", body)
                if not c or c.group(1) == "4294967295":
                    continue
                sid = c.group(1)
                emp["defense_platforms"][sid] = emp["defense_platforms"].get(sid, 0) + 1
        if any(sz and sz.startswith("starbase_") for sz in sizes):
            continue
        if any(sz in STATION_SIZES for sz in sizes):
            for _sid, sz, body in rows:
                kind = STATION_SIZES.get(sz)
                if not kind:
                    continue
                c = re.search(r"coordinate=\s*\{([^}]*)\}", body)
                if not c:
                    continue
                x = re.search(r"x=([-\d.]+)", c.group(1))
                y = re.search(r"y=([-\d.]+)", c.group(1))
                orig = re.search(r"origin=(\d+)", c.group(1))
                if not (x and y and orig):
                    continue
                if kind == "observe":
                    pid = _nearest_planet(
                        float(x.group(1)), float(y.group(1)), orig.group(1), planet_xy, mode="observe"
                    )
                    if pid and pid not in emp["obs_planets"]:
                        emp["obs_planets"].append(pid)
                    continue
                want = "mine" if kind == "mining" else "res"
                pid = _nearest_planet(
                    float(x.group(1)), float(y.group(1)), orig.group(1), planet_xy, want=want
                )
                if not pid:
                    continue
                if str(pid) in colonized and kind != "mining":
                    continue
                key = "mine_planets" if kind == "mining" else "research_planets"
                if pid not in emp[key]:
                    emp[key].append(pid)
            continue
        loc = None
        if rows:
            c = re.search(r"origin=(\d+)", rows[0][2])
            if c:
                loc = c.group(1)
        mil = []
        civ = []
        for _sid, sz, body in rows:
            nm = _ship_display_name(body)
            if sz in MILITARY_SIZES:
                mil.append({"size": sz, "name": nm})
            elif sz in CIVILIAN_SIZES:
                if sz == "transport":
                    continue
                civ.append({"size": sz, "name": nm, "sys": loc})
        fb = fleets.get(fid, "")
        fname = _fleet_display_name(fb)
        if mil:
            emp["named_military"].append({"name": fname, "sys": loc, "ships": mil})
        for sh in civ:
            emp["civilian_fleets"].append(sh)

    # starbase fittings
    mgr = _brace_section(data, "starbase_mgr")
    inner = mgr
    if "starbases=" in mgr[:40]:
        inner = _brace_section("\nstarbases=\n{" + mgr, "starbases") or mgr
    sbs = _split_entries(inner if inner.startswith("\n") else "\n" + inner, tab="\t\t")
    g = _split_entries(_brace_section(data, "galactic_object"), tab="\t")
    sb_sys = {}
    for sid, body in g.items():
        m = re.search(r"starbases=\s*\{([^}]*)\}", body)
        if not m:
            continue
        for t in m.group(1).split():
            if t.isdigit() and t != "4294967295":
                sb_sys[t] = str(sid)
    for sbid, blk in sbs.items():
        st = re.search(r"\n\t\t\tstation=(\d+)", blk) or re.search(r"station=(\d+)", blk)
        station = st.group(1) if st else ""
        # station= is a ship id; owned_fleets lists fleet ids
        fleet_id = None
        if station in fleet_ships:
            fleet_id = station
        else:
            for fid, rows in fleet_ships.items():
                if any(sid == station for sid, _sz, _b in rows):
                    fleet_id = fid
                    break
        owner = fleet_owner.get(fleet_id or station)
        sys_id = sb_sys.get(sbid)
        if not sys_id and fleet_id in fleet_ships:
            c = re.search(r"origin=(\d+)", fleet_ships[fleet_id][0][2])
            if c:
                sys_id = c.group(1)
        if owner not in by_id or not sys_id:
            continue
        mods = re.search(r"modules=\s*\{([^}]*)\}", blk)
        blds = re.search(r"buildings=\s*\{([^}]*)\}", blk)
        by_id[owner]["sb_fit"][sys_id] = {
            "modules": _slot_map(mods.group(1) if mods else ""),
            "buildings": _slot_map(blds.group(1) if blds else ""),
        }

    megas = _split_entries(_brace_section(data, "megastructures"), tab="\t")
    for _mid, body in megas.items():
        typ = re.search(r'type="([^"]+)"', body)
        own = re.search(r"\n\t\towner=(\d+)", body)
        orig = re.search(r"origin=(\d+)", body)
        if not (typ and own and orig):
            continue
        t = typ.group(1)
        if t in SKIP_MEGA_TYPES or "ruined" in t:
            continue
        if own.group(1) in by_id:
            by_id[own.group(1)]["mega_owned"].append({"type": t, "sys": orig.group(1)})

    leaders = _split_entries(_brace_section(data, "leaders"), tab="\t")
    for cid, body in csec.items():
        if cid not in by_id:
            continue
        m = re.search(r"owned_leaders=\s*\{([^}]*)\}", body)
        if not m:
            continue
        got = []
        for lid in m.group(1).split():
            lb = leaders.get(lid, "")
            if not lb:
                continue
            cls = re.search(r'class="([^"]+)"', lb)
            if not cls or cls.group(1) not in LEADER_CLASSES:
                continue
            keys = re.findall(r'key="([^"]+)"', lb[:800])
            raw_name = name_from_keys(keys) if keys else None
            if raw_name in ("Unknown", "") or (raw_name and raw_name.startswith("%")):
                raw_name = None
            age = re.search(r"\n\t\tage=(\d+)", lb)
            skill = re.search(r"\n\t\tlevel=(\d+)", lb)
            traits = [
                t for t in re.findall(r'\n\t\ttraits="([^"]+)"', lb)
                if t.startswith(("leader_trait_", "trait_", "subclass_"))
            ]
            gender = re.search(r'\ngender=([A-Za-z]+)', lb)
            role = "pool"
            loc = re.search(r"location=\s*\{([^}]*)\}", lb)
            loc_s = loc.group(1) if loc else ""
            if "type=council_position" in loc_s:
                role = "councilor"
            elif "type=planet" in loc_s:
                role = "governor"
            elif cls.group(1) == "scientist" and ("type=ship" in loc_s or "type=fleet" in loc_s):
                role = "scientist"
            elif cls.group(1) == "commander" and "type=army" in loc_s:
                role = "general"
            elif cls.group(1) == "commander":
                # 4.x admirals sit on fleet (type=fleet / type=2), not only ship.
                role = "admiral"
            elif cls.group(1) == "official" and not got:
                role = "ruler"
            got.append({
                "name": raw_name,
                "class": cls.group(1),
                "age": int(age.group(1)) if age else 30,
                "skill": max(1, int(skill.group(1))) if skill else 1,
                "traits": traits[:10],
                "role": role,
                "gender": gender.group(1) if gender else None,
            })
            if len(got) >= 10:
                break
        if got and not any(x["role"] == "ruler" for x in got):
            for x in got:
                if x["class"] == "official":
                    x["role"] = "ruler"
                    break
        by_id[cid]["leaders"] = got

    districts = _split_entries(_brace_section(data, "districts"), tab="\t")
    if not districts:
        districts = _split_entries(_brace_section(data, "districts"), tab="")
    zones = _split_entries(_brace_section(data, "zones"), tab="\t")
    buildings = _split_entries(_brace_section(data, "buildings"), tab="\t")
    cols = _split_entries(_brace_section(data, "colony"), tab="\t")
    c2p = {}
    for col_id, body in cols.items():
        m = re.search(r"carrier=\s*\{\s*type=planet\s*reference=(\d+)", body)
        if m:
            c2p[col_id] = m.group(1)

    def _btype(bid):
        b = buildings.get(str(bid), "")
        m = re.search(r'type="([^"]+)"', b)
        return m.group(1) if m else None

    def _zinfo(zid):
        if str(zid) == "4294967295":
            return None
        z = zones.get(str(zid), "")
        if not z:
            return None
        t = re.search(r'type="([^"]+)"', z)
        bids = re.search(r"buildings=\s*\{([^}]*)\}", z)
        blist = []
        if bids:
            for tok in bids.group(1).split():
                if tok.isdigit() and tok != "4294967295":
                    bt = _btype(tok)
                    if bt:
                        blist.append(bt)
        return {"type": t.group(1) if t else "zone_default", "buildings": blist}

    for c in countries:
        layouts = {}
        for col in c.get("owned_planets") or []:
            pid = c2p.get(str(col))
            blk = cols.get(str(col), "")
            if not pid or not blk:
                continue
            dlist = re.search(r"districts=\s*\{([^}]*)\}", blk)
            if not dlist:
                continue
            slots = []
            for did in dlist.group(1).split():
                if not did.isdigit():
                    continue
                d = districts.get(did, "")
                dt = re.search(r'type="([^"]+)"', d)
                lv = re.search(r"level=(\d+)", d)
                zids = re.search(r"zones=\s*\{([^}]*)\}", d)
                zinfos = []
                if zids:
                    for zid in zids.group(1).split():
                        info = _zinfo(zid)
                        if info:
                            zinfos.append(info)
                if dt:
                    slots.append({
                        "type": dt.group(1),
                        "level": max(1, int(lv.group(1))) if lv else 1,
                        "zones": zinfos,
                    })
            if slots:
                layouts[str(pid)] = slots
        c["colony_layout"] = layouts

    feds = {}
    for fid, fbody in _split_entries(_brace_section(data, "federation"), tab="\t").items():
        m = re.search(r"members=\s*\{([^}]*)\}", fbody)
        ids = [t for t in (m.group(1).split() if m else []) if t.isdigit() and t != "4294967295"]
        feds[fid] = ids
    cid_to_fed = {}
    for fid, ids in feds.items():
        for mid in ids:
            cid_to_fed[mid] = fid
    for c in countries:
        fid = cid_to_fed.get(c["id"]) or c.get("_federation")
        ids = feds.get(str(fid), []) if fid is not None else []
        c["federation_peers"] = [x for x in ids if x != c["id"]]

    army_rows = {}
    for aid, body in _split_entries(_brace_section(data, "army"), tab="\t").items():
        own = re.search(r"\n\t\towner=(\d+)", body)
        typ = re.search(r'type="([^"]+)"', body)
        col = re.search(r"\n\t\tcolony=(\d+)", body)
        nm = re.search(r'key="fmt"\s*value=\s*\{\s*key="([^"]+)"', body)
        name = nm.group(1) if nm and not nm.group(1).startswith("%") else None
        planet = c2p.get(col.group(1)) if col else None
        ship = re.search(r"\n\t\tship=(\d+)", body)
        ship_id = ship.group(1) if ship and ship.group(1) != "4294967295" else None
        army_rows[aid] = {
            "owner": own.group(1) if own else None,
            "type": typ.group(1) if typ else "assault_army",
            "planet": planet,
            "ship": ship_id,
            "name": name,
        }
    for col_id, body in cols.items():
        pid = c2p.get(str(col_id))
        if not pid:
            continue
        am = re.search(r"\n\t\tarmy=\s*\{([^}]*)\}", body)
        if not am:
            continue
        for aid in am.group(1).split():
            if aid in army_rows and not army_rows[aid].get("planet") and not army_rows[aid].get("ship"):
                army_rows[aid]["planet"] = pid
    by_owner_armies = {}
    for rec in army_rows.values():
        own = rec.get("owner")
        if own not in by_id:
            continue
        by_owner_armies.setdefault(own, []).append(rec)
    for cid, recs in by_owner_armies.items():
        emp = by_id[cid]
        planet_ars = [a for a in recs if a.get("planet") and not a.get("ship")]
        embarked = [a for a in recs if a.get("ship")]
        leftover = [a for a in recs if not a.get("planet") and not a.get("ship")]
        out = [{"planet": a["planet"], "type": a["type"], "name": a.get("name")} for a in planet_ars]
        emp["armies"] = out
        ship_to_fleet = {}
        fleet_origin = {}
        for fid, rows in fleet_ships.items():
            for sid, _sz, body in rows:
                ship_to_fleet[sid] = fid
                orig = re.search(r"origin=(\d+)", body)
                if orig and orig.group(1) != "4294967295":
                    fleet_origin[fid] = orig.group(1)
        by_fl = {}
        for a in embarked:
            fid = ship_to_fleet.get(a["ship"], a["ship"])
            by_fl.setdefault(fid, []).append(a["type"])
        emp["embarked_fleets"] = [
            {"types": types, "sys": fleet_origin.get(fid)}
            for fid, types in by_fl.items()
        ]
        emp["embarked_armies"] = [{"type": t} for types in by_fl.values() for t in types]

    sites = []
    sites_blob = _nested_section(_brace_section(data, "archaeological_sites"), "sites")
    site_blocks = _split_entries(sites_blob, tab="\t\t") if sites_blob else {}
    if not site_blocks:
        site_blocks = _split_entries(_brace_section(data, "archaeological_sites"), tab="\t")
    for _sid, body in site_blocks.items():
        typ = re.search(r'\ntype="([^"]+)"', body) or re.search(r'type="([^"]+)"', body)
        loc = re.search(r"location=\s*\{([^}]*)\}", body)
        pid = None
        if loc:
            pid_m = re.search(r"\bid=(\d+)", loc.group(1))
            if pid_m:
                pid = pid_m.group(1)
        if not pid:
            pl = re.search(r"planet=(\d+)", body)
            if pl:
                pid = pl.group(1)
        stype = typ.group(1) if typ else None
        if pid and _valid_site_type(stype):
            sites.append({"planet": pid, "type": stype})
    for c in countries:
        c["sites"] = sites
    return countries


def tag_cp_flags(countries, prefix, flags, planet_flags, idx_of):
    for emp in countries:
        idx = idx_of.get(emp["id"])
        if idx is None:
            continue
        for sid in (emp.get("starbases") or {}):
            flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_s{sid}")
        for sid, size in (emp.get("dsc") or {}).items():
            flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_s{sid}")
            if str(size).endswith("_1"):
                short = "dsc1"
            elif str(size).endswith("_2"):
                short = "dsc2"
            else:
                short = "dsc"
            flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_{short}")
        for sid in (emp.get("system_fleets") or {}):
            flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_s{sid}")
        for mega in emp.get("mega_owned") or []:
            flags.setdefault(str(mega["sys"]), []).append(f"{prefix}_{idx}_s{mega['sys']}")
        for i, pid in enumerate(emp.get("mine_planets") or []):
            planet_flags.setdefault(str(pid), []).append(f"{prefix}_{idx}_mine_{i}")
        for i, pid in enumerate(emp.get("research_planets") or []):
            planet_flags.setdefault(str(pid), []).append(f"{prefix}_{idx}_res_{i}")
        for i, pid in enumerate(emp.get("obs_planets") or []):
            planet_flags.setdefault(str(pid), []).append(f"{prefix}_{idx}_obs_{i}")
        for i, site in enumerate(emp.get("sites") or []):
            fl = f"continuum_site_{i}"
            lst = planet_flags.setdefault(str(site["planet"]), [])
            if fl not in lst:
                lst.append(fl)
        for sh in emp.get("civilian_fleets") or []:
            if sh.get("sys"):
                flags.setdefault(str(sh["sys"]), []).append(f"{prefix}_{idx}_s{sh['sys']}")
        for fl in emp.get("named_military") or []:
            if fl.get("sys"):
                flags.setdefault(str(fl["sys"]), []).append(f"{prefix}_{idx}_s{fl['sys']}")
        for fl in emp.get("embarked_fleets") or []:
            if fl.get("sys"):
                flags.setdefault(str(fl["sys"]), []).append(f"{prefix}_{idx}_s{fl['sys']}")
        for sid in (emp.get("defense_platforms") or {}):
            flags.setdefault(str(sid), []).append(f"{prefix}_{idx}_s{sid}")


def wrap_event(eid, flag, inner, kind="event"):
    if not inner.strip():
        inner = "			# none this slice\n"
    return f"""{kind} = {{
	id = {eid}
	is_triggered_only = yes
	hide_window = yes
	immediate = {{
		if = {{
			limit = {{ has_global_flag = {flag} }}
		}}
		else = {{
			set_global_flag = {flag}
{inner}		}}
	}}
}}

"""


def emit_pop_event(restored):
    parts = []
    for r in restored:
        parts.append(r["pop_txt"])
    return wrap_event("continuum_empire.2", "continuum_emp_pops_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_starbase_event(restored):
    parts = []
    for r in restored:
        parts.append(r["sb_txt"])
    return wrap_event("continuum_empire.3", "continuum_emp_starbases_done", "\n".join(parts) + ("\n" if parts else ""))


def _named_fleet_block(emp, cap_flag, owner, prefix=None, idx=None):
    """One create_fleet per Pre military fleet, at that fleet's system, with its name and mix."""
    mil = emp.get("named_military") or []
    if not mil:
        return None
    fe_map = emp.get("_fe_size_designs") or {}
    gfx = emp.get("graphical_culture") or ""
    gfx_bit = ""
    if emp.get("type") == "fallen_empire" and gfx and re.match(r"^[A-Za-z][A-Za-z0-9_]*$", str(gfx)):
        gfx_bit = f" graphical_culture = {gfx}"
    raw_pre = resolved_ship_prefix(emp) or str(emp.get("ship_prefix") or "").replace('"', "").strip()
    if (not raw_pre) or raw_pre.startswith("%") or raw_pre.startswith("PRESCRIPTED_"):
        opref = "O"
    elif raw_pre.startswith("O") and len(raw_pre) > 1:
        opref = raw_pre
    else:
        opref = "O" + raw_pre
    fleets = []
    for fl in mil:
        ships_all = list(fl.get("ships") or [])
        if not ships_all:
            continue
        sys_id = fl.get("sys")
        if prefix is not None and idx is not None and sys_id and str(sys_id) != "4294967295":
            sys_flag = f"{prefix}_{idx}_s{sys_id}"
        else:
            sys_flag = cap_flag
        nline = _fleet_name_line(fl.get("name"))
        for chunk_i in range(0, len(ships_all), 40):
            ships = ships_all[chunk_i:chunk_i + 40]
            creates = []
            for sh in ships:
                sz = sh.get("size") or "corvette"
                named = fe_map.get(sz)
                hull = f'design = "{named}"' if named else f"random_existing_design = {sz}"
                nm = sh.get("name")
                if nm and _is_format_name(nm):
                    nm = None
                nml = f"name = {ship_name_token(nm)}" if nm else "name = random"
                pretty = pretty_ship_name(nm) if nm else None
                old_nm = f'"{opref} {pretty}"' if pretty else "random"
                creates.append(
                    "								if = {\n"
                    "									limit = { owner = { has_country_flag = continuum_old_twin } }\n"
                    f"									create_ship = {{ name = {old_nm} {hull} prefix = no{gfx_bit} }}\n"
                    "								}\n"
                    "								else = {\n"
                    f"									create_ship = {{ {nml} {hull} prefix = yes{gfx_bit} }}\n"
                    "								}"
                )
            inner = "\n".join(creates)
            fleets.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {sys_flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							{nline}
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
{inner}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    return "\n".join(fleets) if fleets else None


def _civilian_blocks(emp, prefix, idx, owner):
    parts = []
    for i, sh in enumerate(emp.get("civilian_fleets") or []):
        sz = sh.get("size") or "science"
        nm = sh.get("name")
        if nm and _is_format_name(nm):
            nm = None
        sys_id = sh.get("sys")
        if not sys_id or str(sys_id) == "4294967295":
            sys_flag = f"{prefix}_{idx}_capital"
        else:
            sys_flag = f"{prefix}_{idx}_s{sys_id}"
        nml = f"name = {ship_name_token(nm)}" if nm else "name = random"
        parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {sys_flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								create_ship = {{ {nml} random_existing_design = {sz} prefix = yes }}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    return "\n".join(parts)


def emit_fleet_event(restored, fallback_fleet_fn):
    parts = []
    for r in restored:
        is_fe = r["emp"].get("type") == "fallen_empire" or str(r.get("prefix") or "").startswith("continuum_fe")
        named = _named_fleet_block(r["emp"], r["cap_flag"], r["owner"], r.get("prefix"), r.get("idx"))
        if named:
            parts.append(named)
        elif not is_fe:
            parts.append(fallback_fleet_fn(r["emp"], r["cap_flag"], r["owner"]))
        civ = _civilian_blocks(r["emp"], r["prefix"], r["idx"], r["owner"])
        if civ:
            parts.append(civ)
    return wrap_event("continuum_empire.4", "continuum_emp_fleets_done", "\n".join(p for p in parts if p) + "\n")


def emit_station_event(restored):
    parts = []
    for r in restored:
        owner = r["owner"]
        prefix = r["prefix"]
        idx = r["idx"]
        for i, _pid in enumerate(r["emp"].get("mine_planets") or []):
            parts.append(f"""			every_galaxy_planet = {{
				limit = {{
					has_planet_flag = {prefix}_{idx}_mine_{i}
					NOT = {{ exists = mining_station }}
				}}
				if = {{
					limit = {{ exists = event_target:{owner} }}
					create_mining_station = {{ owner = event_target:{owner} }}
				}}
			}}""")
        for i, _pid in enumerate(r["emp"].get("research_planets") or []):
            parts.append(f"""			every_galaxy_planet = {{
				limit = {{
					has_planet_flag = {prefix}_{idx}_res_{i}
					NOT = {{ exists = research_station }}
				}}
				if = {{
					limit = {{ exists = event_target:{owner} }}
					create_research_station = {{ owner = event_target:{owner} }}
				}}
			}}""")
        for i, _pid in enumerate(r["emp"].get("obs_planets") or []):
            parts.append(f"""			every_galaxy_planet = {{
				limit = {{ has_planet_flag = {prefix}_{idx}_obs_{i} }}
				if = {{
					limit = {{ exists = event_target:{owner} }}
					create_fleet = {{
						settings = {{ spawn_debris = no }}
						effect = {{
							set_owner = event_target:{owner}
							create_ship = {{ name = "Observation Post" random_existing_design = observation_station prefix = no }}
							set_location = prev
						}}
					}}
				}}
			}}""")
        plats = r["emp"].get("defense_platforms") or {}
        if plats:
            parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				event_target:{owner} = {{
					give_technology = {{ tech = tech_space_defense_station_1 message = no }}
					refresh_auto_generated_ship_designs = yes
				}}
			}}""")
        for sid, n in plats.items():
            n = int(n)
            if n < 1:
                continue
            parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {prefix}_{idx}_s{sid} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
								while = {{
									count = {n}
									create_ship = {{ name = random random_existing_design = military_station_small }}
								}}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    return wrap_event("continuum_empire.5", "continuum_emp_stations_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_module_mega_event(restored):
    parts = []
    for r in restored:
        owner = r["owner"]
        prefix = r["prefix"]
        idx = r["idx"]
        emp = r["emp"]
        for sid, fit in (emp.get("sb_fit") or {}).items():
            mods = []
            for slot, mod in sorted((fit.get("modules") or {}).items()):
                mods.append(f"					set_starbase_module = {{ slot = {int(slot) + 1} module = {mod} }}")
            for slot, bld in sorted((fit.get("buildings") or {}).items()):
                mods.append(f"					set_starbase_building = {{ slot = {int(slot) + 1} building = {bld} }}")
            if not mods:
                continue
            inner = "\n".join(mods)
            parts.append(f"""			every_system = {{
				limit = {{
					has_star_flag = {prefix}_{idx}_s{sid}
					exists = starbase
				}}
				starbase = {{
{inner}
				}}
			}}""")
        for mega in emp.get("mega_owned") or []:
            parts.append(f"""			every_megastructure = {{
				limit = {{
					is_megastructure_type = {mega['type']}
					solar_system = {{ has_star_flag = {prefix}_{idx}_s{mega['sys']} }}
				}}
				if = {{
					limit = {{ exists = event_target:{owner} }}
					set_owner = event_target:{owner}
				}}
			}}""")
    return wrap_event("continuum_empire.6", "continuum_emp_modules_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_leader_event(restored):
    parts = []
    for r in restored:
        owner = r["owner"]
        emp = r["emp"]
        leads = emp.get("leaders") or []
        if not leads:
            continue
        creates = [
            "					every_owned_leader = {",
            "						kill_leader = { show_notification = no }",
            "					}",
        ]
        for n, ld in enumerate(leads):
            valid = [
                t for t in (ld.get("traits") or [])
                if t and re.match(r"^[A-Za-z][A-Za-z0-9_]*$", t)
            ]
            at_create = [t for t in valid if not _is_councilor_trait(t)]
            deferred = [t for t in valid if _is_councilor_trait(t)]
            trait_block = ("\n						traits = { " + " ".join(f"trait = {t}" for t in at_create) + " }") if at_create else ""
            nm = "random" if not ld.get("name") else _script_token(ld["name"])
            gender = ld.get("gender")
            gender_line = ""
            if gender in ("female", "male", "indeterminate"):
                gender_line = f"\n						gender = {gender}"
            tgt = f"{owner}_ld_{n}"
            creates.append(f"""					create_leader = {{
						name = {nm}
						species = owner_main_species
						class = {ld['class']}
						skill = {int(ld['skill'])}
						set_age = {int(ld['age'])}
						randomize_traits = no{gender_line}{trait_block}
						effect = {{
							save_event_target_as = {tgt}
							set_leader_flag = {tgt}
						}}
					}}""")
            role = ld["role"]
            if role == "ruler":
                creates.append(f"					assign_leader = event_target:{tgt}")
            elif role == "governor":
                creates.append(f"""					capital_scope = {{
						assign_leader = event_target:{tgt}
					}}""")
            elif role == "admiral":
                creates.append(f"""					random_owned_fleet = {{
						limit = {{
							NOT = {{ exists = leader }}
							any_owned_ship = {{
								OR = {{
									is_ship_size = corvette
									is_ship_size = frigate
									is_ship_size = destroyer
									is_ship_size = cruiser
									is_ship_size = battleship
									is_ship_size = titan
									is_ship_size = small_ship_fallen_empire
									is_ship_size = large_ship_fallen_empire
									is_ship_size = massive_ship_fallen_empire
								}}
							}}
						}}
						assign_leader = event_target:{tgt}
					}}""")
            elif role == "scientist":
                # assign_leader is country/fleet/planet/army — not ship.
                creates.append(f"""					random_owned_fleet = {{
						limit = {{
							NOT = {{ exists = leader }}
							any_owned_ship = {{ is_ship_size = science }}
						}}
						assign_leader = event_target:{tgt}
					}}""")
            if deferred and role == "ruler":
                adds = "\n".join(
                    f"						add_trait = {{ trait = {t} show_message = no }}"
                    for t in deferred
                )
                creates.append(f"""					event_target:{tgt} = {{
{adds}
					}}""")
        inner = "\n".join(creates)
        parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				event_target:{owner} = {{
{inner}
				}}
			}}""")
    return wrap_event("continuum_empire.7", "continuum_emp_leaders_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_diplo_event(restored):
    owners = [r["owner"] for r in restored]
    parts = []
    for i, a in enumerate(owners):
        for b in owners[i + 1:]:
            parts.append(f"""			if = {{
				limit = {{
					exists = event_target:{a}
					exists = event_target:{b}
				}}
				event_target:{a} = {{
					establish_communications_no_message = event_target:{b}
				}}
			}}""")
    id_to_owner = {r["emp"]["id"]: r["owner"] for r in restored}
    for r in restored:
        ov = r["emp"].get("overlord")
        if ov and ov in id_to_owner and id_to_owner[ov] != r["owner"]:
            parts.append(f"""			if = {{
				limit = {{
					exists = event_target:{r['owner']}
					exists = event_target:{id_to_owner[ov]}
				}}
				event_target:{r['owner']} = {{
					set_subject_of = {{ who = event_target:{id_to_owner[ov]} }}
				}}
			}}""")
    seen_fed = set()
    for r in restored:
        a = r["owner"]
        for pid in r["emp"].get("federation_peers") or []:
            if pid not in id_to_owner:
                continue
            b = id_to_owner[pid]
            if a == b:
                continue
            pair = tuple(sorted((a, b)))
            if pair in seen_fed:
                continue
            seen_fed.add(pair)
            parts.append(f"""			if = {{
				limit = {{
					exists = event_target:{pair[0]}
					exists = event_target:{pair[1]}
				}}
				event_target:{pair[0]} = {{
					join_alliance = {{ who = event_target:{pair[1]} override_requirements = yes }}
				}}
			}}""")
    return wrap_event("continuum_empire.8", "continuum_emp_diplo_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_intel_event(restored):
    # Canary Pre has no usable country-intel block; keep the event so we can fill it later.
    return wrap_event("continuum_empire.9", "continuum_emp_intel_done", "")


def emit_layout_event(restored):
    parts = []
    for r in restored:
        owner = r["owner"]
        prefix = r["prefix"]
        idx = r["idx"]
        emp = r["emp"]
        pops = list((emp.get("colony_pop") or {}).keys())
        layouts = emp.get("colony_layout") or {}
        gestalt = _is_gestalt(emp)
        for i, pid in enumerate(pops):
            slots = layouts.get(str(pid))
            if not slots:
                continue
            lines = [
                "					set_planet_flag = ignore_ai_building_limitations",
                "					remove_all_buildings = yes",
                "					remove_all_districts = yes",
            ]
            for slot in slots:
                dt = slot["type"]
                zones = slot.get("zones") or []
                # 4.x city district is one slot with N zones, not N separate districts.
                repeats = 1 if zones else min(int(slot.get("level") or 1), 8)
                for _ in range(repeats):
                    lines.append(f"					add_district = {{ district_type = {dt} ignore_cap = yes }}")
                for zi, z in enumerate(zones):
                    lines.append(
                        f"					add_zone = {{ district = {dt} zone = {z['type']} zone_slot = {zi} replace = yes }}"
                    )
                    for bld in z.get("buildings") or []:
                        if bld in ("building_colony_shelter", "building_colony_shelter_machine", "building_colony_shelter_hive"):
                            continue
                        if bld == "building_holo_theatres" and gestalt:
                            continue
                        if str(bld).startswith("building_galactic_memorial"):
                            continue
                        if bld == "building_factory_1":
                            if gestalt:
                                continue
                            fdist = dt if dt == "district_industrial" else "district_city"
                            lines.append(
                                f"					add_zone = {{ district = {fdist} zone = zone_industrial zone_slot = {zi} replace = yes }}"
                            )
                            lines.append(
                                f"					add_building = {{ district = {fdist} zone = zone_industrial building = building_factory_1 }}"
                            )
                            continue
                        lines.append(
                            f"					add_building = {{ district = {dt} zone = {z['type']} building = {bld} }}"
                        )
            if gestalt or emp.get("type") == "fallen_empire":
                lines.append("					remove_building = building_colony_shelter")
            else:
                lines.append("					validate_and_repair_planet_buildings_and_districts = yes")
            inner = "\n".join(lines)
            parts.append(f"""			every_galaxy_planet = {{
				limit = {{
					has_planet_flag = {prefix}_{idx}_c{i}
					is_colony = yes
					exists = owner
					is_owned_by = event_target:{owner}
				}}
{inner}
			}}""")
    return wrap_event("continuum_empire.10", "continuum_emp_layouts_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_army_event(restored, eid="continuum_empire.11", flag="continuum_emp_armies_done", kind="event", include_embarked=True):
    parts = []
    for r in restored:
        owner = r["owner"]
        prefix = r["prefix"]
        idx = r["idx"]
        emp = r["emp"]
        e_fleets = emp.get("embarked_fleets") or []
        if include_embarked and not e_fleets and emp.get("embarked_armies"):
            e_fleets = [{"types": [a.get("type") for a in emp.get("embarked_armies") or []]}]
        wipe_types = []
        seen_w = set()
        for fl in e_fleets:
            for t in fl.get("types") or []:
                typ = _army_type_key(t)
                if typ in seen_w or _is_job_defense_army(typ):
                    continue
                seen_w.add(typ)
                wipe_types.append(typ)
        if wipe_types:
            or_bits = " ".join(f"army_type = {t}" for t in wipe_types)
            parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				event_target:{owner} = {{
					every_owned_army = {{
						limit = {{ OR = {{ {or_bits} }} }}
						remove_army = yes
					}}
				}}
			}}""")
        pops = list((emp.get("colony_pop") or {}).keys())
        pid_to_c = {str(p): i for i, p in enumerate(pops)}
        by_c = {}
        for ar in emp.get("armies") or []:
            if _is_job_defense_army(ar.get("type")):
                continue
            ci = pid_to_c.get(str(ar.get("planet") or ""))
            if ci is None:
                continue
            by_c.setdefault(ci, []).append(ar)
        for ci, ars in by_c.items():
            creates = []
            for ar in ars:
                typ = _army_type_key(ar.get("type"))
                creates.append(
                    f"				create_army = {{ owner = event_target:{owner} species = owner_main_species type = {typ} }}"
                )
            if not creates:
                continue
            inner = "\n".join(creates)
            parts.append(f"""			every_galaxy_planet = {{
				limit = {{
					has_planet_flag = {prefix}_{idx}_c{ci}
					is_colony = yes
					exists = owner
					is_owned_by = event_target:{owner}
				}}
{inner}
			}}""")
        if include_embarked:
            cap_flag = r.get("cap_flag") or f"{prefix}_{idx}_capital"
            for fl in e_fleets:
                types = [_army_type_key(t) for t in (fl.get("types") or []) if t]
                if not types:
                    continue
                sys_id = fl.get("sys")
                if sys_id and str(sys_id) != "4294967295":
                    sys_flag = f"{prefix}_{idx}_s{sys_id}"
                else:
                    sys_flag = cap_flag
                inner = "\n".join(
                    f"							create_army_transport = {{ army_type = {typ} species = owner_main_species }}"
                    for typ in types
                )
                parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				random_system = {{
					limit = {{ has_star_flag = {sys_flag} }}
					event_target:{owner} = {{
						create_fleet = {{
							settings = {{ spawn_debris = no }}
							effect = {{
								set_owner = prev
{inner}
								set_location = prevprev
							}}
						}}
					}}
				}}
			}}""")
    return wrap_event(eid, flag, "\n".join(parts) + ("\n" if parts else ""), kind=kind)


def emit_site_event(restored):
    parts = []
    seen = set()
    sites = []
    for r in restored:
        for s in r["emp"].get("sites") or []:
            key = (str(s.get("planet")), s.get("type"))
            if not key[0] or not _valid_site_type(key[1]) or key in seen:
                continue
            seen.add(key)
            sites.append(s)
    for i, s in enumerate(sites):
        parts.append(f"""			every_galaxy_planet = {{
				limit = {{ has_planet_flag = continuum_site_{i} }}
				create_archaeological_site = {s['type']}
			}}""")
    return wrap_event("continuum_empire.12", "continuum_emp_sites_done", "\n".join(parts) + ("\n" if parts else ""))


def emit_tradition_event(restored):
    parts = []
    for r in restored:
        owner = r["owner"]
        trads = [t for t in (r["emp"].get("traditions") or []) if _valid_tradition(t)][:20]
        if not trads:
            continue
        inner = "\n".join(f"					add_tradition = {t}" for t in trads)
        parts.append(f"""			if = {{
				limit = {{ exists = event_target:{owner} }}
				event_target:{owner} = {{
{inner}
				}}
			}}""")
    return wrap_event("continuum_empire.13", "continuum_emp_traditions_done", "\n".join(parts) + ("\n" if parts else ""))


def cp_summary(emp):
    return (
        f"mine={len(emp.get('mine_planets') or [])} "
        f"res={len(emp.get('research_planets') or [])} "
        f"obs={len(emp.get('obs_planets') or [])} "
        f"ships={sum(len(f.get('ships') or []) for f in (emp.get('named_military') or []))} "
        f"civ={len(emp.get('civilian_fleets') or [])} "
        f"leaders={len(emp.get('leaders') or [])} "
        f"megas={len(emp.get('mega_owned') or [])} "
        f"layouts={len(emp.get('colony_layout') or {})} "
        f"armies={len(emp.get('armies') or [])} "
        f"trads={len(emp.get('traditions') or [])} "
        f"fed={len(emp.get('federation_peers') or [])} "
        f"sites={len(emp.get('sites') or [])}"
    )
