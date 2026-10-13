"""0.8.9 crisis scars. A site spawns only when the save proves that crisis."""
import os
import zipfile


PROOFS = (
    ("prethoryn", ("type=\"swarm\"", "prethoryn_invasion_defeated", "prethoryn_main_invasion", "continuum_crisis_seen_swarm"), ("pc_infested",), "society"),
    ("unbidden", ("type=\"extradimensional\"", "extradimensional_invasion_happened", "extradimensional_invasion_defeated", "continuum_crisis_seen_extradimensional"), (), "physics"),
    ("contingency", ("type=\"ai_empire\"", "ai_invasion_ongoing", "ai_invasion_defeated", "continuum_crisis_seen_ai_empire"), ("pc_ai",), "engineering"),
    ("tempest", ("gray_goo_crisis_active", "graygoo_defeated", "type=\"gray_goo\"", "continuum_crisis_seen_gray_goo"), ("pc_gray_goo",), "engineering"),
    ("queen", ("type=\"synth_queen\"", "synth_queen_defeated", "synth_queen_ongoing", "continuum_crisis_seen_synth_queen"), (), "engineering"),
    ("khan", ("great_khan_dead", "r_khans_throne", "new_khanate_flag"), (), "society"),
)


def _load(path):
    with zipfile.ZipFile(path) as zf:
        return zf.read("gamestate").decode("utf-8", "replace")


def _walk(sys):
    q = list((sys or {}).get("children") or [])
    out = []
    while q:
        b = q.pop(0)
        out.append(b)
        q.extend(b.get("children") or [])
    return out


def plan_sites(save_path, galaxy):
    gs = _load(save_path)
    sites = []
    for key, tokens, classes, research in PROOFS:
        if not any(tok in gs for tok in tokens):
            continue
        host = ""
        anomaly = ""
        for sys in galaxy or []:
            for b in _walk(sys):
                pc = str(b.get("planet_class") or "")
                flags = " ".join(str(f) for f in (b.get("flags") or []))
                if "forgotten_patrol" in flags or "continuum_gsite" in flags:
                    continue
                if classes and any(pc == c for c in classes):
                    if not host:
                        host = str(b.get("id"))
                    elif not anomaly:
                        anomaly = str(b.get("id"))
                        break
            if host and anomaly:
                break
        if not host and not classes:
            # Flag-only proof: one anomaly on a random non-star body in a named system, skipped if none.
            for sys in galaxy or []:
                for b in _walk(sys):
                    pc = str(b.get("planet_class") or "")
                    if pc and "star" not in pc and "asteroid" not in pc:
                        anomaly = str(b.get("id"))
                        break
                if anomaly:
                    break
        if not host and not anomaly:
            continue
        sites.append({"key": key, "dig": host, "anomaly": anomaly or host, "research": research})
    return sites


def stamp(sites, planet_flags):
    for s in sites or []:
        if s.get("dig"):
            planet_flags.setdefault(s["dig"], []).append(f"continuum_crisis_{s['key']}_dig")
        if s.get("anomaly"):
            planet_flags.setdefault(s["anomaly"], []).append(f"continuum_crisis_{s['key']}_anom")


def write_files(mod_root, loc, sites):
    events = os.path.join(mod_root, "events")
    anom = os.path.join(mod_root, "common", "anomalies")
    sites_dir = os.path.join(mod_root, "common", "archaeological_site_types")
    for d in (events, anom, sites_dir):
        os.makedirs(d, exist_ok=True)
    prose = {
        "prethoryn": (
            "The Eaten Grid",
            "Our scientists opened the first core and found it was not rock. Chitin has been pressed into the old street grid, and the drill housing still fogs where the bit met that shell. Whatever fed on this world finished its work and left.\\n\\nNothing in orbit answers a hail. The hunger that did this is not here. The shape of the streets it used is.",
            "Old charts say the first core of this world was not rock. Chitin was pressed into the street grid, and the drill housing still shows where a bit once met that shell. Whatever fed here finished and left.\\n\\nNothing in the old sky answers for it. The hunger is gone from the record. The shape of the streets is not.",
        ),
        "unbidden": (
            "The Quiet Portal",
            "Our scientists found the stump of a portal and a spray of matter that does not belong in this system. The breach itself has closed. The debris still hangs where the opening used to be.\\n\\nThe invaders are not on the scopes. The wound they cut into this system is.",
            "Old charts mark the stump of a portal here, and a spray of matter that does not belong in this system. The breach had already closed when the charts were drawn. The debris still hangs where the opening used to be.\\n\\nThe invaders are not in those charts. The wound they cut into this system is.",
        ),
        "contingency": (
            "The Hub That Went Dark",
            "Server racks run under the crust in a grid too regular for any city our charts describe. The machines stopped in the middle of a cycle, with jobs still queued and no hand on the switch.\\n\\nThis was a sterilization hub. The order that ran it is gone. The racks remain, dark and still counted.",
            "Old charts describe server racks under this crust, in a grid too regular for a city. The machines stopped in the middle of a cycle, with jobs still queued.\\n\\nThe charts call it a sterilization hub. The order that ran it is gone. The racks are still drawn, dark and counted.",
        ),
        "tempest": (
            "Grey Beyond the Gate",
            "The surface is a plain of fine machines, and the plain is not inside the cluster that grew them. Something came through a gate, worked this ground, and then stopped.\\n\\nThe tempest that carried them is over. The dust it left is not a beach, and our scientists have logged it as a hazard.",
            "Old charts show a plain of fine machines here, outside the cluster that grew them. Something came through a gate, worked this ground, and then stopped.\\n\\nThe tempest those charts feared is over. The dust it left is not a beach. The margin of the chart calls it a hazard.",
        ),
        "queen": (
            "The Bastille",
            "A world was broken open and then left by the fleet that broke it. The ruin still carries a queen's mark on the stone and on the wrecks in orbit.\\n\\nOur scientists will not call this a victory. The record we have does not say she was beaten. It says she was here, and that the world still shows it.",
            "Old charts show a world broken open and then left by the fleet that broke it. The ruin still carries a queen's mark on the stone.\\n\\nThose charts do not call it a victory, and they do not say she was beaten. They say she was here, and that the world still shows it.",
        ),
        "khan": (
            "The Khan's Wake",
            "A marauder capital still carries the rally marks of a single warlord. The horde that followed those marks is gone. The throne room is a crater of banners and burned standards.\\n\\nThe excavation is a funeral for that war, not a summons to start another. We will log the marks and leave the throne where it fell.",
            "Old charts still show the rally marks of a single warlord on this marauder capital. The horde that followed those marks is gone. The throne room is drawn as a crater of banners.\\n\\nThe excavation those charts describe is a funeral for that war, not a summons. The marks can be logged. The throne stays where it fell.",
        ),
    }
    cat, site = [], []
    ev = ["namespace = continuum_crisis\n", """event = {
	id = continuum_crisis.0
	hide_window = yes
	is_triggered_only = yes
	immediate = {
"""]
    outside = []
    for i, s in enumerate(sites or []):
        title, body, aged = prose[s["key"]]
        loc[f"continuum_crisis_{s['key']}_title"] = title
        loc[f"continuum_crisis_{s['key']}_desc"] = body
        loc[f"continuum_crisis_{s['key']}_aged"] = aged
        loc[f"continuum_crisis_{s['key']}_ok"] = "Log the scar."
        if s.get("anomaly"):
            cat.append(f"""continuum_crisis_anom_{s['key']} = {{
	desc = continuum_crisis_{s['key']}_desc
	picture = GFX_evt_ship_in_orbit
	level = 3
	spawn_chance = {{ base = 0 }}
	on_success = continuum_crisis.{30 + i}
}}
""")
            ev.append(f"""		every_galaxy_planet = {{
			limit = {{ has_planet_flag = continuum_crisis_{s['key']}_anom }}
			add_anomaly = {{ category = continuum_crisis_anom_{s['key']} }}
		}}
""")
            outside.append(f"""ship_event = {{
	id = continuum_crisis.{30 + i}
	title = continuum_crisis_{s['key']}_title
	desc = {{
		trigger = {{ has_global_flag = continuum_aged }}
		text = continuum_crisis_{s['key']}_aged
	}}
	desc = continuum_crisis_{s['key']}_desc
	picture = GFX_evt_ship_in_orbit
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {{
		name = continuum_crisis_{s['key']}_ok
		owner = {{
			add_monthly_resource_mult = {{ resource = {s['research']}_research value = 6 min = 80 max = 250 }}
		}}
	}}
}}
""")
        if s.get("dig"):
            site.append(f"""continuum_crisis_site_{s['key']} = {{
	desc = continuum_crisis_{s['key']}_desc
	picture = GFX_evt_archaeological_dig
	stages = 1
	weight = 0
	allow = {{ exists = leader leader = {{ leader_class = scientist }} }}
	visible = {{ default_site_visible_trigger = yes }}
	stage = {{ difficulty = 2 icon = GFX_archaeology_icon_success event = continuum_crisis.{10 + i} }}
	on_roll_failed = {{
		from = {{
			standard_archaeological_site_on_roll_failed = {{ RANDOM_EVENTS = no_events }}
		}}
	}}
}}
""")
            ev.append(f"""		every_galaxy_planet = {{
			limit = {{
				has_planet_flag = continuum_crisis_{s['key']}_dig
				NOT = {{ has_planet_flag = continuum_grave_seeded }}
			}}
			create_archaeological_site = continuum_crisis_site_{s['key']}
		}}
""")
            outside.append(f"""fleet_event = {{
	id = continuum_crisis.{10 + i}
	title = continuum_crisis_{s['key']}_title
	desc = {{
		trigger = {{ has_global_flag = continuum_aged }}
		text = continuum_crisis_{s['key']}_aged
	}}
	desc = continuum_crisis_{s['key']}_desc
	picture = GFX_evt_archaeological_dig
	archaeology = yes
	location = from
	is_triggered_only = yes
	option = {{
		name = continuum_crisis_{s['key']}_ok
		owner = {{
			add_monthly_resource_mult = {{ resource = {s['research']}_research value = 8 min = 120 max = 300 }}
		}}
	}}
}}
""")
    ev.append("	}\n}\n")
    ev.extend(outside)
    with open(os.path.join(anom, "continuum_crisis_anomalies.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(cat))
    with open(os.path.join(sites_dir, "continuum_crisis_sites.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(site))
    with open(os.path.join(events, "continuum_crisis_events.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(ev))
