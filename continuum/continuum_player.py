"""Intro + copy-of-Pre handling. 4.4 titles cannot use trigger/text blocks."""

WELCOME = "Welcome to your galaxy...continued!"


def pretty_empire_name(name):
    import continuum_cp
    name = name or "Empire"
    if name.upper().startswith("EMPIRE DESIGN"):
        rest = name[13:].strip(" :_-")
        return rest[:1].upper() + rest[1:] if rest else "Empire"
    if " " in name and not name.startswith(("SPEC_", "NAME_", "PRESCRIPTED_", "EMPIRE_DESIGN_", "%")):
        return name
    return continuum_cp.display_key(name) or name


def pretty_leader_name(name):
    import continuum_cp
    if not name or str(name).startswith("%"):
        return None
    name = str(name)
    if name.startswith("councilor_"):
        return None
    looked = continuum_cp.loc_lookup(name)
    if looked:
        return looked
    if "_CHR_" in name:
        return name.split("_CHR_", 1)[-1].replace("_", " ")
    if name.startswith(("SPEC_", "NAME_", "PRESCRIPTED_", "EMPIRE_DESIGN_")):
        return continuum_cp.display_key(name)
    return name


def _welcome(rest):
    return f"{WELCOME}\\n\\n{rest}"


def _ok(token, prefix):
    return bool(token) and str(token).isidentifier() and str(token).startswith(prefix)


def copy_match_lines(emp, sp_class):
    """Triggers that mean the player designed the same government as this Pre empire."""
    lines = []
    if sp_class and str(sp_class).isidentifier():
        lines.append(f"is_species_class = {sp_class}")
    auth = emp.get("authority") or ""
    if _ok(auth, "auth_"):
        lines.append(f"has_authority = {auth}")
    civics = [c for c in (emp.get("civics") or []) if _ok(c, "civic_")]
    ethics = [e for e in (emp.get("ethics") or []) if _ok(e, "ethic_")]
    if len(civics) < 1 or len(ethics) < 1:
        return None
    for c in civics:
        lines.append(f"has_civic = {c}")
    for e in ethics:
        lines.append(f"has_ethic = {e}")
    origin = emp.get("origin") or ""
    if _ok(origin, "origin_") and origin != "origin_default":
        lines.append(f"has_origin = {origin}")
    return lines


def _tab(lines, n):
    pre = "\t" * n
    return "\n".join(pre + ln for ln in lines)


def _ship_prefix_old(raw):
    """O + original prefix, e.g. HMS -> OHMS."""
    s = str(raw or "").replace('"', "").replace("\n", "").strip()
    if not s or s.startswith("%") or s.startswith("PRESCRIPTED_"):
        return "O"
    if s.startswith("O") and len(s) > 1:
        return s
    return "O" + s


_SIMILAR_COLOR = {
    "red": "burgundy", "burgundy": "dark_red", "dark_red": "red",
    "orange": "dark_orange", "dark_orange": "orange", "light_orange": "orange",
    "yellow": "orange", "bright_yellow": "yellow",
    "blue": "dark_blue", "dark_blue": "indigo", "light_blue": "blue",
    "indigo": "dark_purple", "turquoise": "teal", "teal": "dark_teal",
    "dark_teal": "green", "green": "dark_green", "dark_green": "green",
    "light_green": "green", "purple": "dark_purple", "dark_purple": "purple",
    "pink": "burgundy", "black": "dark_grey", "dark_grey": "grey",
    "grey": "dark_grey", "light_grey": "grey", "white": "light_grey",
    "beige": "brown", "brown": "dark_brown", "dark_brown": "brown",
}


def _color_token(c):
    s = str(c or "null").strip().strip('"')
    if not s or s == "null":
        return "null"
    if s.isidentifier():
        return s
    return "black"


def _similar_color(c):
    s = _color_token(c)
    if s == "null":
        return "dark_grey"
    alt = _SIMILAR_COLOR.get(s)
    if alt:
        return alt
    return "dark_grey" if s != "dark_grey" else "black"


def reverse_flag_colors(colors):
    """Swap primary/secondary; if that does nothing, shift to a similar color."""
    cols = [_color_token(x) for x in (colors or [])]
    while len(cols) < 4:
        cols.append("null")
    a, b, ic, extra = cols[0], cols[1], cols[2], cols[3]
    if b not in ("null", a):
        a, b = b, a
    else:
        a = _similar_color(a)
        if a == b:
            a = _similar_color(a)
    if ic not in ("null", "") and ic in (a, b):
        ic = _similar_color(ic)
        if ic in (a, b):
            ic = "white" if "white" not in (a, b) else "black"
    return [a, b, ic, extra]


def _flag_file(val, fallback):
    s = str(val or fallback).replace('"', "").replace("\n", "").strip()
    return s or fallback


def _set_shared_loc(loc_entries):
    loc_entries["continuum_intro_title"] = "Continuum Present"
    loc_entries["continuum_intro_ok"] = "Begin"
    loc_entries["continuum_intro_present"] = _welcome(
        "Only a few years have passed. The old empires still hold their space. "
        "You are a new political entity in their galaxy — not their heir."
    )
    loc_entries["continuum_intro_new"] = _welcome(
        "Your people have reached FTL. The other empires still occupy their cores. "
        "The gaps between them are yours to claim."
    )
    loc_entries["continuum_intro_border"] = _welcome(
        "You rise on a frontier between the old empires. They still exist. "
        "You are a new state on their doorstep."
    )
    loc_entries["continuum_intro_primitive"] = _welcome(
        "Your world was pre-FTL. You have reached the stars. "
        "The empires that mapped this sky still exist. You do not serve them."
    )
    loc_entries["continuum_intro_crisis"] = _welcome(
        "Only a few years have passed since the crisis that scarred this galaxy. "
        "The old empires that survived still hold their cores. You are a new polity in the wreckage."
    )
    loc_entries["continuum_intro_copy"] = _welcome(
        "Another empire already uses this government. They still hold their space. "
        "You are a new polity under the same banner — not their heir. You keep your name."
    )


def _copy_cases(restored, loc_entries):
    """Per-empire copy matches. 4.4 has no has_name, so government is the stand-in."""
    _set_shared_loc(loc_entries)
    cases = []
    for r in restored:
        if r.get("prefix") != "continuum_emp":
            continue
        emp = r["emp"]
        lines = copy_match_lines(emp, r.get("_sp_class") or "")
        if not lines:
            continue
        idx = r["idx"]
        owner = r["owner"]
        pretty = pretty_empire_name(emp.get("name") or "Empire")
        old_key = f"continuum_old_emp_{idx}"
        loc_entries[old_key] = f"Old {pretty}"
        loc_key = f"continuum_intro_copy_{idx}"
        loc_entries[loc_key] = _welcome(
            f"The {pretty} still hold their space. You rise as a new polity under the same banner — "
            "not their heir. You keep your name."
        )
        leader_locs = []
        for n, ld in enumerate(emp.get("leaders") or []):
            pretty_ld = pretty_leader_name(ld.get("name"))
            if not pretty_ld:
                continue
            lk = f"continuum_c_ld_{idx}_{n}"
            loc_entries[lk] = f"O. {pretty_ld}"
            leader_locs.append((n, lk))
        cases.append({
            "loc_key": loc_key,
            "lines": lines,
            "idx": idx,
            "owner": owner,
            "old_key": old_key,
            "ship_prefix": _ship_prefix_old(emp.get("ship_prefix")),
            "leader_locs": leader_locs,
            "flag_icon_cat": _flag_file(emp.get("flag_icon_cat"), "special"),
            "flag_icon": _flag_file(emp.get("flag_icon"), "pirate_flag.dds"),
            "flag_bg_cat": _flag_file(emp.get("flag_bg_cat"), "backgrounds"),
            "flag_bg": _flag_file(emp.get("flag_bg"), "00_solid.dds"),
            "flag_colors": reverse_flag_colors(emp.get("flag_colors")),
        })
    return cases


def _copy_rename_block(case, leaders=True, match_from=None):
    owner = case["owner"]
    leader_ifs = []
    if leaders:
        for n, lk in case["leader_locs"]:
            leader_ifs.append(
                "				every_owned_leader = {\n"
                f"					limit = {{ has_leader_flag = {owner}_ld_{n} }}\n"
                f"					set_name = {lk}\n"
                "				}"
            )
    leaders_txt = ("\n" + "\n".join(leader_ifs)) if leader_ifs else ""
    cols = case["flag_colors"]
    color_txt = " ".join(f'"{c}"' for c in cols)
    if match_from:
        limit = (
            "			limit = {\n"
            f"				exists = event_target:{match_from}\n"
            f"				exists = event_target:{owner}\n"
            f"				event_target:{match_from} = {{\n"
            f"{_tab(case['lines'], 5)}\n"
            "				}\n"
            "			}\n"
        )
    else:
        limit = (
            "			limit = {\n"
            "				exists = event_target:" + owner + "\n"
            f"{_tab(case['lines'], 4)}\n"
            "			}\n"
        )
    return (
        "		if = {\n"
        f"{limit}"
        f"			event_target:{owner} = {{\n"
        f"				set_name = {case['old_key']}\n"
        f"				set_country_flag = continuum_old_twin\n"
        f"				set_ship_prefix = \"{case['ship_prefix']}\"\n"
        "				change_country_flag = {\n"
        f"					icon = {{ category = \"{case['flag_icon_cat']}\" file = \"{case['flag_icon']}\" }}\n"
        f"					background = {{ category = \"{case['flag_bg_cat']}\" file = \"{case['flag_bg']}\" }}\n"
        f"					colors = {{ {color_txt} }}\n"
        "				}\n"
        f"{leaders_txt}\n"
        "			}\n"
        "		}"
    )


def emit_copy_identity_effects(restored, loc_entries):
    """Name / prefix / flag in empire.1, scoped to the player so fleets spawn as OHMS."""
    cases = _copy_cases(restored, loc_entries)
    txt = "\n".join(
        _copy_rename_block(c, leaders=False, match_from="continuum_human") for c in cases
    )
    return (txt + "\n") if txt.strip() else "		# no Pre government copy\n"


def emit_player_event(restored, loc_entries):
    """Hidden: Old-prefix, C ship prefix, C. leaders, reversed flag when player copies."""
    cases = _copy_cases(restored, loc_entries)
    rename_txt = "\n".join(_copy_rename_block(c) for c in cases)
    if not rename_txt.strip():
        rename_txt = "		# no Pre government matches"
    return f"""namespace = continuum_player
country_event = {{
	id = continuum_player.1
	is_triggered_only = yes
	hide_window = yes
	trigger = {{ is_ai = no }}
	immediate = {{
{rename_txt}
	}}
}}
"""


def emit_intro_event(restored, loc_entries, had_crisis=False):
    cases = _copy_cases(restored, loc_entries)
    desc_copy = []
    for case in cases:
        lines = case["lines"]
        loc_key = case["loc_key"]
        desc_copy.append(
            "	desc = {\n"
            "		trigger = {\n"
            f"{_tab(lines, 3)}\n"
            "		}\n"
            f"		text = {loc_key}\n"
            "	}"
        )
    desc_copy_txt = ("\n".join(desc_copy) + "\n") if desc_copy else ""
    crisis_flag = "			set_global_flag = continuum_had_crisis\n" if had_crisis else ""
    return f"""namespace = continuum_intro
country_event = {{
	id = continuum_intro.1
	is_triggered_only = yes
	title = continuum_intro_title
{desc_copy_txt}	desc = {{
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
"""
