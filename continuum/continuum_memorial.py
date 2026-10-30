"""Aged dig sites for human empires the weed removed.

A normal anomaly is cleared when a colony is founded, and a new pre-FTL
may settle the old capital. A dig site stays on that world. Present still
has the empire, so the site is seeded only under continuum_aged.
"""
import os

import continuum_empires as ce

_AUTH = {
    "auth_democratic": "democracy",
    "auth_oligarchic": "oligarchy",
    "auth_dictatorial": "dictatorship",
    "auth_imperial": "empire",
    "auth_hive_mind": "hive",
    "auth_machine_intelligence": "machine intelligence",
    "auth_corporate": "corporation",
}
_GRAVE = {
    "pc_cracked": "The capital world is cracked open.",
    "pc_tomb": "The capital world is a tomb.",
    "pc_shielded": "The capital world is sealed behind a shield.",
    "pc_infested": "The capital world is overgrown.",
    "pc_ai": "The capital world is a machine shell.",
    "pc_gray_goo": "The capital world is a grey plain of machines.",
}


def _bodies(sys):
    root = (sys or {}).get("hierarchy_root") or sys or {}
    out = []
    q = list(root.get("children") or [])
    while q:
        b = q.pop(0)
        out.append(b)
        q.extend(b.get("children") or [])
    return out


def _body_ids(galaxy):
    found = {}
    for sys in galaxy or []:
        sid = str(sys.get("id"))
        for b in _bodies(sys):
            found[str(b.get("id"))] = (sid, b)
    return found


def _ethic_phrase(ethics):
    words = []
    for e in ethics or []:
        if e == "ethic_gestalt_consciousness":
            continue
        words.append(e.replace("ethic_", "").replace("_", " "))
    if not words:
        return ""
    if len(words) == 1:
        return words[0]
    return ", ".join(words[:-1]) + " and " + words[-1]


def _who(emp, species, capital_name):
    name = ce.display_empire_name(emp.get("name") or "This empire")
    sp = (species or {}).get(str(emp.get("founder_species") or "")) or {}
    sp_name = ce.species_display_name(sp, emp)
    auth = _AUTH.get(emp.get("authority") or "", "government")
    ethics = _ethic_phrase(emp.get("ethics"))
    gov = f"a {auth}"
    if ethics:
        gov = f"a {auth} of {ethics}"
    capital = capital_name or "a capital the charts no longer name"
    origin = str(emp.get("origin") or "")
    origin_bit = ""
    if origin.startswith("origin_") and "default" not in origin:
        plain = origin[len("origin_"):].replace("_", " ")
        origin_bit = f" Their origin is recorded as {plain}."
    return (
        f"The {name} was {gov}. The species that founded them was {sp_name}. "
        f"Their capital was {capital}.{origin_bit}"
    )


def _battle_sentence(emp, owned, battles):
    name = ce.display_empire_name(emp.get("name") or "")
    mine = {str(s) for s in (owned or [])}
    rows = [b for b in (battles or []) if str(b.get("system")) in mine]
    if not rows:
        return ""
    row = max(rows, key=lambda b: int(b.get("size") or 0))
    sides = []
    for label in list(row.get("crises") or []) + list(row.get("victims") or []) + list(row.get("killers") or []):
        if label and label not in sides:
            sides.append(label)
    others = [s for s in sides if s != name]
    where = row.get("name") or "one of their systems"
    if len(sides) >= 2:
        fight = f"a fight in {where} between {sides[0]} and {sides[1]}"
    elif others:
        fight = f"a fight in {where} involving {others[0]}"
    elif sides:
        fight = f"a fight in {where} involving {sides[0]}"
    else:
        fight = f"a fight in {where}. The charts do not name the fleets"
    goal = ""
    goals = [str(g).replace("_", " ") for g in (row.get("wargoals") or []) if g]
    if goals:
        goal = f" The war was logged as {goals[0]}."
    side = ""
    if name and name in (row.get("victims") or []):
        side = f" {name} is recorded among the destroyed."
    elif name and name in (row.get("killers") or []):
        side = f" {name} is recorded among those who fired."
    return f"Our scientists recovered a record of {fight}.{goal}{side}"


def _end(emp, body, battle):
    pc = str((body or {}).get("planet_class") or "")
    grave = _GRAVE.get(pc, "")
    bits = [b for b in (grave, battle) if b]
    if not bits:
        return (
            "The later charts list no heir for this empire and give no recorded cause for the silence. "
            "The transponder stops."
        )
    tail = " The later charts name no cause beyond that. The transponder stops."
    return " ".join(bits) + tail


def plan(empire_plan, aged_galaxy, battles):
    extinct = {str(c) for c in (empire_plan.get("aged_extinct_ids") or [])}
    species = empire_plan.get("species") or {}
    c2p = empire_plan.get("colony_to_planet") or {}
    capitals = empire_plan.get("capitals") or {}
    owned = empire_plan.get("owned") or {}
    bodies = _body_ids(aged_galaxy)
    sys_by = {str(s.get("id")): s for s in aged_galaxy or []}
    out = []
    for emp in empire_plan.get("empires") or []:
        if not emp.get("played_by_human"):
            continue
        if str(emp.get("id")) not in extinct:
            continue
        cap_pid = c2p.get(str(emp.get("capital") or ""))
        if not cap_pid:
            pops = list((emp.get("colony_pop") or {}).keys())
            cap_pid = str(pops[0]) if pops else ""
        host = bodies.get(str(cap_pid or ""))
        sid = capitals.get(emp["id"])
        if host:
            sid = host[0]
            body = host[1]
        else:
            sys = sys_by.get(str(sid or ""))
            body = None
            for b in _bodies(sys):
                pc = str(b.get("planet_class") or "")
                if pc and "star" not in pc and "asteroid" not in pc:
                    body = b
                    break
            if body is None:
                print(f"Player memorial skipped: {emp.get('name')} has no remaining body")
                continue
            cap_pid = str(body.get("id"))
        cap_name = (emp.get("colony_names") or {}).get(str(cap_pid)) or (body.get("name") if body else "")
        who = _who(emp, species, ce.display_empire_name(cap_name) if cap_name else "")
        battle = _battle_sentence(emp, owned.get(emp["id"]), battles)
        end = _end(emp, body, battle)
        name = ce.display_empire_name(emp.get("name") or "Empire")
        prim_planets = {
            str(np.get("planet_id"))
            for np in (empire_plan.get("aged_new_prims") or [])
        }
        # A colony clears an anomaly. A dig would compete with the pre-FTL, so
        # the find moves to the star and the pre-FTL keeps the world.
        mode = "star" if str(cap_pid) in prim_planets else "dig"
        out.append({
            "idx": len(out),
            "mode": mode,
            "planet_id": str(cap_pid),
            "system_id": str(sid or ""),
            "name": name,
            "desc": who + "\\n\\n" + end,
        })
    labels = ", ".join(m["name"] for m in out) if out else "none"
    print(f"Player memorials: {len(out)}" + (f" ({labels})" if out else ""))
    return out


def stamp(memorials, planet_flags, star_flags):
    for m in memorials or []:
        i = m["idx"]
        if m.get("mode") == "star" and m.get("system_id"):
            star_flags.setdefault(str(m["system_id"]), []).append(f"continuum_pmem_star_{i}")
        else:
            planet_flags.setdefault(str(m["planet_id"]), []).append(
                f"continuum_player_memorial_{i}"
            )


def write_files(mod_root, loc, memorials):
    if not memorials:
        return
    events = os.path.join(mod_root, "events")
    sites = os.path.join(mod_root, "common", "archaeological_site_types")
    os.makedirs(events, exist_ok=True)
    os.makedirs(sites, exist_ok=True)
    site_bits = []
    star_cats = []
    ev = ["""namespace = continuum_memorial

event = {
	id = continuum_memorial.0
	hide_window = yes
	is_triggered_only = yes
	trigger = { has_global_flag = continuum_aged }
	immediate = {
"""]
    outside = []
    for m in memorials:
        i = m["idx"]
        key = f"continuum_player_memorial_{i}"
        site_key = f"continuum_pmem_{i}"
        title = f"{m['name']} Last Transponder"
        loc[key] = title
        loc[site_key] = title
        loc[f"{key}_desc"] = m["desc"]
        loc[f"{key}_ok"] = "Log the name."
        if m.get("mode") == "star":
            star_key = f"continuum_pmem_star_{i}"
            loc[star_key] = title
            loc[f"{star_key}_desc"] = m["desc"]
            loc[f"{star_key}_ok"] = "Log the name."
            star_cats.append(f"""{star_key} = {{
	desc = {star_key}_desc
	picture = GFX_evt_ship_in_orbit
	level = 2
	spawn_chance = {{ base = 0 }}
	on_success = continuum_memorial.{20 + i}
}}
""")
            ev.append(f"""		every_system = {{
			limit = {{ has_star_flag = {star_key} }}
			random_system_planet = {{
				limit = {{ is_star = yes }}
				add_anomaly = {{ category = {star_key} }}
			}}
		}}
""")
            outside.append(f"""ship_event = {{
	id = continuum_memorial.{20 + i}
	title = {star_key}
	desc = {star_key}_desc
	picture = GFX_evt_ship_in_orbit
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {{
		name = {key}_ok
		owner = {{
			add_monthly_resource_mult = {{ resource = society_research value = 6 min = 80 max = 250 }}
		}}
	}}
}}
""")
            continue
        site_bits.append(f"""{site_key} = {{
	desc = {key}_desc
	picture = GFX_evt_archaeological_dig
	stages = 1
	weight = 0
	allow = {{
		exists = leader
		leader = {{ leader_class = scientist }}
	}}
	visible = {{
		default_site_visible_trigger = yes
	}}
	stage = {{
		difficulty = 2
		icon = GFX_archaeology_icon_success
		event = continuum_memorial.{10 + i}
	}}
	on_roll_failed = {{
		from = {{
			standard_archaeological_site_on_roll_failed = {{ RANDOM_EVENTS = no_events }}
		}}
	}}
}}
""")
        ev.append(f"""		every_galaxy_planet = {{
			limit = {{ has_planet_flag = {key} }}
			create_archaeological_site = {site_key}
		}}
""")
        outside.append(f"""fleet_event = {{
	id = continuum_memorial.{10 + i}
	title = {key}
	desc = {key}_desc
	picture = GFX_evt_archaeological_dig
	show_sound = event_ship_bridge
	archaeology = yes
	location = from
	is_triggered_only = yes
	option = {{
		name = {key}_ok
		owner = {{
			add_monthly_resource_mult = {{ resource = society_research value = 6 min = 80 max = 250 }}
		}}
	}}
}}
""")
    ev.append("	}\n}\n")
    ev.extend(outside)
    with open(os.path.join(sites, "continuum_player_memorials.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(site_bits))
    if star_cats:
        anom = os.path.join(mod_root, "common", "anomalies")
        os.makedirs(anom, exist_ok=True)
        with open(os.path.join(anom, "continuum_player_memorials.txt"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(star_cats))
    with open(os.path.join(events, "continuum_memorial_events.txt"), "w", encoding="utf-8") as fh:
        fh.write("".join(ev))
