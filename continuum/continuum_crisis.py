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
            "Our scientists opened the first geological core and found it was not composed of rock. Hardened chitin has been pressed into the old street grid, and the drill housing still fogs up where the bit met that biological shell. Whatever massive organism fed on this world finished its grim work and left the system.\n\nnothing in orbit answers a standard hail. The horrific hunger that did this is no longer here. The shape of the streets it used to feed certainly is.",
            "Old geological charts state that the first core extracted from this world was not rock. Dense chitin was pressed directly into the street grid, and the preserved drill housing still shows where a bit once met that organic shell. Whatever fed here finished its meal and departed.\n\nNothing in the old, empty sky answers for it. The hunger itself is gone from the record. The digested shape of the streets is not.",
        ),
        "unbidden": (
            "The Quiet Portal",
            "Our scientists have located the severed stump of a dimensional portal and a massive spray of exotic matter that does not belong in this system's physics. The extradimensional breach itself has closed. The spatial debris still hangs silently where the tearing opening used to be.\n\nThe energy invaders are not currently showing on the scopes. The terrible wound they cut into this star system is.",
            "Old astrometric charts mark the scarred stump of a dimensional portal here, alongside a spray of exotic matter that defies local physics. The breach had apparently already closed when the charts were originally drawn. The spatial debris still hangs silently where the opening used to be.\n\nThe invaders are not mentioned anywhere else in those charts. The deep wound they cut into this system certainly is.",
        ),
        "contingency": (
            "The Hub That Went Dark",
            "Massive server racks run deep under the planetary crust in a grid far too regular for any organic city our charts can describe. The machines abruptly stopped in the middle of a processing cycle, with active jobs still queued in the memory and no physical hand anywhere near the kill switch.\n\nThis was an active sterilization hub. The mechanical order that ran it is gone. The server racks remain, dark and counted.",
            "Old exploratory charts describe vast server racks buried under this crust, arranged in a grid far too regular for an organic city. The machines abruptly stopped right in the middle of a processing cycle, with heavy tasks still sitting queued in the memory.\n\nThe old charts refer to it as a sterilization hub. The controlling order that ran it is long gone. The racks are still accurately drawn on the maps, dark and counted.",
        ),
        "tempest": (
            "Grey Beyond the Gate",
            "The surface below is a sprawling plain of fine, microscopic machines, and the plain is located outside the distant cluster that originally grew them. Something flooded through an L-Gate, worked this ground into dust, and then abruptly stopped its task.\n\nThe nanite tempest that carried them is over. The grey dust it left behind is not a beach, and our scientists have logged it as a severe hazard.",
            "Old survey charts show a massive plain of fine, microscopic machines here, sitting far outside the cluster that originally grew them. Something flooded through an L-Gate, worked this ground into dust, and then simply stopped.\n\nThe devastating tempest those ancient charts feared is over. The grey dust it left behind is definitely not a beach. The scribbled margin of the chart plainly calls it a hazard.",
        ),
        "queen": (
            "The Bastille",
            "A populated world was broken open and then left behind by the fleet that broke it. The sprawling ruin still carries a synthetic queen's distinct mark on the stone and on the shattered wrecks drifting in orbit.\n\nOur scientists will not call this a victory. The official record we have does not say she was ever beaten. It only says she was here, and that the ruined world still shows it.",
            "Old historical charts show a world that was broken open and then left behind by the massive fleet that broke it. The sprawling ruin still clearly carries a synthetic queen's mark stamped into the stone.\n\nThose ancient charts do not call it a victory, and they do not say she was ever beaten. They only say she was here, and that the broken world still shows it.",
        ),
        "khan": (
            "The Khan's Wake",
            "A battered marauder capital still carries the war rally marks of a single, unifying warlord. The massive horde that followed those marks is now gone. The central throne room is nothing but a scorched crater of burned banners and shattered standards.\n\nThis excavation is a quiet funeral for that war, not a loud summons to start another one. We will document the rally marks and leave the broken throne where it fell.",
            "Old cultural charts still show the aggressive rally marks of a single warlord painted across this marauder capital. The sprawling horde that followed those marks is gone. The central throne room is drawn simply as a scorched crater of burned banners.\n\nThe excavation those charts describe is a quiet funeral for that war, not a summons for a new one. The marks can be logged for the archives. The broken throne stays where it fell.",
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
