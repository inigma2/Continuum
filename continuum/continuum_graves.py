"""0.8.6 Graves: visible Pre scars become anomalies and digs on Present and Aged."""
import os

MAX_TOTAL = 18
MAX_PER_KIND = 5

PLANET_KIND = {
    "pc_shattered": "cracked",
    "pc_shattered_2": "cracked",
    "pc_broken": "cracked",
    "pc_cracked": "cracked",
    "pc_shielded": "shielded",
    "pc_ringworld_shielded": "shielded",
    "pc_habitat_shielded": "shielded",
    "pc_nuked": "tomb",
    "pc_infested": "infested",
    "pc_ai": "machine",
    "pc_machine_broken": "machine",
    "pc_shrouded": "shroud",
    "pc_gray_goo": "nanite",
}

KIND_ORDER = ("cracked", "shielded", "infested", "machine", "shroud", "nanite", "tomb", "wronghole")
KIND_NUM = {k: i + 1 for i, k in enumerate(KIND_ORDER)}

KIND_META = {
    "cracked": {"picture": "GFX_evt_exploding_planet", "icon": "GFX_archaeology_runes_E1"},
    "shielded": {"picture": "GFX_evt_in_the_dark", "icon": "GFX_archaeology_runes_B1"},
    "tomb": {"picture": "GFX_evt_tomb_world", "icon": "GFX_archaeology_runes_D1"},
    "infested": {"picture": "GFX_evt_infested", "icon": "GFX_archaeology_runes_C1"},
    "machine": {"picture": "GFX_evt_dead_city", "icon": "GFX_archaeology_runes_A1"},
    "shroud": {"picture": "GFX_evt_psionics", "icon": "GFX_archaeology_runes_F1"},
    "nanite": {"picture": "GFX_evt_gray_goo", "icon": "GFX_archaeology_runes_E2"},
    "wronghole": {"picture": "GFX_evt_black_hole", "icon": "GFX_archaeology_runes_D2"},
}


def _iter_bodies(system):
    root = system.get("hierarchy_root") or system
    stack = list(root.get("children") or [])
    while stack:
        body = stack.pop()
        yield body
        stack.extend(body.get("children") or [])


def _is_star_body(body):
    pc = str(body.get("planet_class") or "")
    return body.get("body_type") == "star" or pc.endswith("_star") or "black_hole" in pc


def _system_flags(system):
    return [str(f) for f in (system.get("flags") or [])]


def classify_graves(galaxy_data):
    """Pick a capped set of visible catastrophe leftovers."""
    buckets = {k: [] for k in KIND_ORDER}
    for sys in galaxy_data or []:
        sid = str(sys.get("id") or "")
        fl = _system_flags(sys)
        if "formless_system" in fl:
            continue
        star_class = str(sys.get("system_star_class") or "")
        bodies = list(_iter_bodies(sys))
        n_planets = sum(1 for b in bodies if not _is_star_body(b))
        has_war_world = False
        for body in bodies:
            if _is_star_body(body):
                continue
            pc = str(body.get("planet_class") or "")
            kind = PLANET_KIND.get(pc)
            if not kind:
                continue
            pid = str(body.get("id") or "")
            if not pid:
                continue
            has_war_world = has_war_world or kind in ("cracked", "tomb", "infested", "machine")
            buckets[kind].append({"planet": pid, "system": sid, "kind": kind, "class": pc})
        if "black_hole" in star_class:
            suspicious = has_war_world or n_planets >= 5
            if suspicious:
                star = next((b for b in bodies if _is_star_body(b)), None)
                pid = str((star or {}).get("id") or "")
                if pid:
                    buckets["wronghole"].append(
                        {"planet": pid, "system": sid, "kind": "wronghole", "class": star_class}
                    )
    picked = []
    for kind in KIND_ORDER:
        rows = buckets[kind][:MAX_PER_KIND]
        picked.extend(rows)
        if len(picked) >= MAX_TOTAL:
            picked = picked[:MAX_TOTAL]
            break
    for i, g in enumerate(picked):
        g["idx"] = i
    return picked


def apply_grave_flags(graves, planet_flags):
    flags = planet_flags if planet_flags is not None else {}
    for g in graves:
        lst = flags.setdefault(str(g["planet"]), [])
        kind_flag = f"continuum_grave_{g['kind']}"
        site_flag = f"continuum_gsite_{g['idx']}"
        if kind_flag not in lst:
            lst.append(kind_flag)
        if site_flag not in lst:
            lst.append(site_flag)
    return flags


KIND_SITE_NAME = {
    "cracked": "Fault-City",
    "shielded": "Pacifier Ring",
    "tomb": "Glass District",
    "infested": "Dry Nest",
    "machine": "Edge Node",
    "shroud": "The Shore of Maybe",
    "nanite": "Quiet Shore",
    "wronghole": "The Witness Orbit",
}


def write_grave_files(mod_root, loc_entries, graves):
    """Static script + loc. Safe to call with an empty graves list."""
    kinds = KIND_ORDER
    common = os.path.join(mod_root, "common")
    os.makedirs(os.path.join(common, "anomalies"), exist_ok=True)
    os.makedirs(os.path.join(common, "archaeological_site_types"), exist_ok=True)
    os.makedirs(os.path.join(mod_root, "events"), exist_ok=True)
    _write_anomaly_categories(os.path.join(common, "anomalies", "continuum_graves.txt"), kinds)
    _write_site_types(os.path.join(common, "archaeological_site_types", "continuum_graves.txt"), graves)
    _write_events(os.path.join(mod_root, "events", "continuum_graves_events.txt"), kinds, graves)
    loc_entries.update(GRAVE_LOC)
    for g in graves:
        loc_entries[f"continuum_gsite_{g['idx']}"] = KIND_SITE_NAME.get(g["kind"], "Grave")
        loc_entries[f"continuum_gsite_{g['idx']}_desc"] = GRAVE_LOC.get(
            f"continuum_site_{g['kind']}_desc", "A scar from the old galaxy."
        )
        loc_entries[f"continuum_gsite_{g['idx']}_desc_aged"] = GRAVE_LOC.get(
            f"continuum_site_{g['kind']}_desc_aged", "A scar from a prior age."
        )
    counts = {}
    for g in graves:
        counts[g["kind"]] = counts.get(g["kind"], 0) + 1
    print("Graves:", ", ".join(f"{k}={n}" for k, n in counts.items()) if counts else "none")
    if graves:
        print(f"  unique dig types: {len(graves)}")


def _write_anomaly_categories(path, kinds):
    lines = ["# Continuum 0.8.6 - visible catastrophe leftovers. spawn_chance 0: seeded by event.\n"]
    for kind in kinds:
        pic = KIND_META[kind]["picture"]
        lines.append(f"""continuum_grave_{kind}_cat = {{
	desc = continuum_grave_{kind}_cat_desc
	picture = {pic}
	level = 3
	spawn_chance = {{
		base = 0
	}}
	on_success = continuum_graves.{KIND_NUM[kind]}
}}
""")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def _write_site_types(path, graves):
    lines = ["# Continuum 0.8.6: one site type per planet (Stellaris treats type as unique).\n"]
    for g in graves:
        kind = g["kind"]
        meta = KIND_META[kind]
        n = KIND_NUM[kind]
        idx = g["idx"]
        lines.append(f"""continuum_gsite_{idx} = {{
	desc = continuum_gsite_{idx}_desc
	picture = {meta["picture"]}
	stages = 3
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
		icon = {meta["icon"]}
		event = continuum_graves.{n * 10 + 1}
	}}
	stage = {{
		difficulty = 3
		icon = {meta["icon"]}
		event = continuum_graves.{n * 10 + 2}
	}}
	stage = {{
		difficulty = 4
		icon = {meta["icon"]}
		event = continuum_graves.{n * 10 + 3}
	}}
	on_roll_failed = {{
		from = {{
			standard_archaeological_site_on_roll_failed = {{ RANDOM_EVENTS = no_events }}
		}}
	}}
}}
""")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def _write_events(path, kinds, graves):
    bits = [
        "namespace = continuum_graves\n",
        """event = {
	id = continuum_graves.0
	hide_window = yes
	is_triggered_only = yes
	immediate = {
""",
    ]
    for g in graves:
        idx = g["idx"]
        bits.append(f"""		every_galaxy_planet = {{
			limit = {{
				has_planet_flag = continuum_gsite_{idx}
				NOT = {{ has_planet_flag = continuum_grave_seeded }}
			}}
			create_archaeological_site = continuum_gsite_{idx}
			set_planet_flag = continuum_grave_seeded
		}}
""")
    bits.append("	}\n}\n")
    for kind in kinds:
        pic = KIND_META[kind]["picture"]
        n = KIND_NUM[kind]
        bits.append(f"""ship_event = {{
	id = continuum_graves.{n}
	title = continuum_grave_{kind}_scan_title
	desc = {{
		trigger = {{ has_global_flag = continuum_aged }}
		text = continuum_grave_{kind}_scan_desc_aged
	}}
	desc = continuum_grave_{kind}_scan_desc
	picture = {pic}
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {{
		name = continuum_grave_scan_ok
		owner = {{
			add_monthly_resource_mult = {{
				resource = society_research
				value = 6
				min = 80
				max = 250
			}}
		}}
	}}
}}
""")
        for stage in (1, 2, 3):
            last = stage == 3
            reward = ""
            if last:
                reward = """
		owner = {
			add_monthly_resource_mult = {
				resource = society_research
				value = 12
				min = 200
				max = 600
			}
			add_resource = { minerals = 200 alloys = 80 }
			if = {
				limit = { has_ancrel = yes }
				small_artifact_reward = yes
			}
		}"""
            bits.append(f"""fleet_event = {{
	id = continuum_graves.{n * 10 + stage}
	title = continuum_grave_{kind}_{stage}_title
	desc = {{
		trigger = {{ has_global_flag = continuum_aged }}
		text = continuum_grave_{kind}_{stage}_desc_aged
	}}
	desc = continuum_grave_{kind}_{stage}_desc
	picture = {pic}
	show_sound = event_wind_ruins
	location = from
	archaeology = yes
	is_triggered_only = yes
	option = {{
		name = continuum_grave_dig_ok{reward}
	}}
}}
""")
    # Console-fireable stories (event continuum_graves.201 etc.) so loc can be
    # tested without a valid archaeology UI.
    for kind in kinds:
        n = KIND_NUM[kind]
        pic = KIND_META[kind]["picture"]
        bits.append(f"""country_event = {{
	id = continuum_graves.{200 + n}
	title = continuum_grave_{kind}_scan_title
	desc = {{
		trigger = {{ has_global_flag = continuum_aged }}
		text = continuum_grave_{kind}_scan_desc_aged
	}}
	desc = continuum_grave_{kind}_scan_desc
	picture = {pic}
	show_sound = event_ship_bridge
	is_triggered_only = yes
	option = {{
		name = continuum_grave_scan_ok
	}}
}}
""")
    bits.append("""country_event = {
	id = continuum_graves.900
	title = continuum_grave_qa_title
	desc = continuum_grave_qa_desc
	picture = GFX_evt_exploding_planet
	show_sound = event_ship_bridge
	location = event_target:continuum_test_grave
	is_triggered_only = yes
	immediate = {
		every_galaxy_planet = {
			limit = { has_planet_flag = continuum_gsite_0 }
			save_event_target_as = continuum_test_grave
			set_surveyed = { surveyed = yes surveyor = root }
			solar_system = {
				save_event_target_as = continuum_test_grave_sys
				set_surveyed = { surveyed = yes surveyor = root }
			}
		}
		set_visited = event_target:continuum_test_grave_sys
		if = {
			limit = {
				any_owned_fleet = {
					any_owned_ship = { is_ship_size = science }
				}
			}
			random_owned_fleet = {
				limit = {
					any_owned_ship = { is_ship_size = science }
				}
				clear_orders = yes
				set_location = {
					target = event_target:continuum_test_grave
					distance = 8
					angle = random
				}
			}
		}
		else = {
			create_fleet = {
				name = continuum_grave_qa_ship
				settings = { spawn_debris = no }
				effect = {
					set_owner = root
					create_ship = {
						name = continuum_grave_qa_ship
						random_existing_design = science
						prefix = no
					}
					set_location = {
						target = event_target:continuum_test_grave
						distance = 8
						angle = random
					}
					save_event_target_as = continuum_test_digger
				}
			}
			create_leader = {
				class = scientist
				species = owner_main_species
				name = random
				skill = 3
			}
			event_target:continuum_test_digger = {
				assign_leader = last_created_leader
			}
		}
	}
	option = {
		name = continuum_grave_qa_ok
	}
}
""")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(bits))


GRAVE_LOC = {
    "continuum_grave_scan_ok": "Log it. We will dig.",
    "continuum_grave_dig_ok": "Continue the excavation",
    "continuum_grave_qa_title": "Fault-City Approach",
    "continuum_grave_qa_desc": "A science ship is in orbit of the first cracked grave. Stay paused. Select the ship in the outliner. Close the planet window. Right-click the world and choose Excavate Archaeological Site.",
    "continuum_grave_qa_ok": "On station",
    "continuum_grave_qa_ship": "Grave Digger",
    # --- cracked ---
    "continuum_grave_cracked_cat": "The Shattered Grid",
    "continuum_grave_cracked_cat_desc": "From orbit the planet looks like a fruit struck by a hammer. The crust has failed along straight lines, not along tectonic grain. Those lines meet at right angles. Cities, or the ghosts of cities, still glitter in the fault walls: glass, alloy, the faint checker of streets that no longer have a surface to rest on. This was not a natural quake. Something in this system learned how to open a world.",
    "continuum_grave_cracked_cat_desc_aged": "Time has sanded the wound, but not enough. The planet is still a cracked bowl of slag and night. In the oldest fault the sensors find a lattice of melted streets, too regular for geology. Whoever lived here is a rumor. Whoever broke them left a signature that has outlasted nations.",
    "continuum_grave_cracked_scan_title": "A World Folded Open",
    "continuum_grave_cracked_scan_desc": "The science officer does not swear. She only enlarges the image until a river of metal is obviously a boulevard, until a cluster of rectangles is obviously a district, until the heat bloom at the center is obviously a single night. Colossus fire, or something that wanted to be one. The atmosphere is a scream that never finished. We can dig the fault. We should.",
    "continuum_grave_cracked_scan_desc_aged": "The first officer calls it a wound. The second calls it a quarry. Both are right. Millennia of vacuum have not hidden the grid. Down in the dark the alloy still remembers being walls. If we listen with the right instruments, this planet will tell us who it used to be, and what kind of hate could teach a starship to crack a world.",
    "continuum_site_cracked": "Fault-City",
    "continuum_site_cracked_desc": "An excavation along the primary fracture, where the last streets still cling to the cliff.",
    "continuum_site_cracked_desc_aged": "A dig in slag that used to be a capital. The layers are a calendar of one terrible hour.",
    "continuum_grave_cracked_1_title": "Street Names in the Slag",
    "continuum_grave_cracked_1_desc": "We lower a probe into a canyon that was a transit spine. The walls are laminated: hab-block, then office, then a transit tube that still holds a train like a fossil. Personal effects have become ore. A child's toy has become a nodule of rare metal. The computers are dead, but the street layout is a language. Someone planned this city. Someone else ended the plan in a single firing solution.",
    "continuum_grave_cracked_1_desc_aged": "The first layer is not dust. It is people, compressed until they are chemistry. We do not find bones so much as alloys that used to be bones' neighbors: kitchen steel, wedding wire, the screws of a million doors. The excavation AI reconstructs a plaza. In the plaza there is a statue plinth with no statue. The inscription is gone. The violence that took the statue also took the name of the species that carved it. We will have to guess them from their plumbing.",
    "continuum_grave_cracked_2_title": "The Firing Solution",
    "continuum_grave_cracked_2_desc": "Deeper in the fault we find glass that is not volcanic. It is crust flash-melted from above, then frozen while it was still raining inward. The melt pattern points at a single vector: a weapon in high orbit, held on target long enough to walk a line across the hemisphere. World-cracker doctrine. Or a crisis weapon wearing the same grin. Either way, this was not an accident and it was not a war of years. It was an afternoon.",
    "continuum_grave_cracked_2_desc_aged": "Scholars of dead wars argue in the logs. One insists this is a Colossus scar: too clean, too linear. Another says no empire would waste a planet-killer on a world this small unless the world had become a symbol. A third, quieter, notes that some crisis fleets learned the same trick without ever asking permission. The slag does not take sides. It only remembers the angle of the beam.",
    "continuum_grave_cracked_3_title": "What the Streets Were For",
    "continuum_grave_cracked_3_desc": "At the bottom of the grid we find a civic core: data vaults, a transit authority, a mural of a sky that no longer exists. We cannot recover a language, but we can recover a mood. They liked parks. They numbered their years. They did not expect the sky to become a weapon. The last packet in the vault is not a plea. It is a maintenance schedule. They thought they had a tomorrow. We take samples, maps, and a silence that will follow the ship for a while.",
    "continuum_grave_cracked_3_desc_aged": "The last chamber is a library that burned so hot it vitrified. In the glass there are still page-shapes, like fish in ice. We will never read them. We can still tell they were loved: the shelves were carved, not stamped. Whoever cracked this world did not merely kill a population. They interrupted a conversation that had been going for centuries. We bag the glass. We log the vector. We leave a beacon so the next curious ship will not have to guess whether this was geology. It was not.",
    # --- shielded ---
    "continuum_grave_shielded_cat": "The Quiet Veil",
    "continuum_grave_shielded_cat_desc": "The planet is intact. That is the horror. Cities sit under a perfect dome of force, streets uncracked, parks still green if the spectrum is to be believed. No impact winter. No craters. Someone wrapped this world in a Global Pacifier and left it like a specimen in a jar. The people inside, if they still think, have been thinking for a very long time without mail.",
    "continuum_grave_shielded_cat_desc_aged": "A jewel in a bottle. The shield has outlasted the empire that poured it. Under the veil the cities have gone dark, or perhaps only quiet. From here it looks like peace. It is the peace of a held breath that was never allowed to finish.",
    "continuum_grave_shielded_scan_title": "Cities in a Bottle",
    "continuum_grave_shielded_scan_desc": "No colossus crack. No firestorm. The megastructure, if you can call a planetary shield a megastructure, is a Global Pacifier: a sentence without an execution. We can study the projector scars in orbit and the cultural layer beneath the veil. We cannot, from here, knock.",
    "continuum_grave_shielded_scan_desc_aged": "Myths say some worlds were 'set aside' by old weapons that preferred imprisonment to ash. This is one of those sentences, still being served. The veil does not answer hails. It does not decay on any clock we trust. Digging the orbital ring of projectors may tell us who had the key, and whether they ever meant to come back.",
    "continuum_site_shielded": "Pacifier Ring",
    "continuum_site_shielded_desc": "Orbital wreckage of the projectors that still hold the world silent.",
    "continuum_site_shielded_desc_aged": "A ring of dead machines keeping a living-looking planet in museum time.",
    "continuum_grave_shielded_1_title": "The Key That Was Not Left",
    "continuum_grave_shielded_1_desc": "The projectors are military. Serials, cooling fins, the ugly elegance of a Colossus support craft. Someone stood in this orbit and chose not to crack the world. That is a kind of mercy and a kind of cruelty. We find no key. We find a checklist with the last box unticked: 'reversal window.' They planned to come back. They did not.",
    "continuum_grave_shielded_1_desc_aged": "Rust, then not-rust: the shield-eaters are still humming in a language of capacitors. Around them the rest of the station has gone to driftwood. A mural on an interior bulkhead shows a planet and a hand covering it, gently. The hand has too many fingers for comfort, or perhaps just enough for the species that painted it. We still do not know if the people under the veil would thank that hand.",
    "continuum_grave_shielded_2_title": "A Weather Report From Nowhere",
    "continuum_grave_shielded_2_desc": "We bounce a passive scan through a thinning lobe of the field. Parks. Transit. A stadium whose lights are off. No fires. If anyone is alive they have learned to live without us, and without the sky. The ethics officer writes three pages and then deletes two of them. There is no protocol for a world that was punished by being preserved.",
    "continuum_grave_shielded_2_desc_aged": "Legend would call them sleepers. The instruments call them a thermal average. If there are people, they are living in a bottle garden with no night that we can grant them. If there are not, then this is the most expensive tomb ever poured. Either way the weapon did what it was designed to do: it stopped a story without ending the planet.",
    "continuum_grave_shielded_3_title": "Mercy, Logged",
    "continuum_grave_shielded_3_desc": "The last log on the projector is not a gloat. It is a date for a hearing that never sat. 'Pacification pending review.' We take the serials, the field-geometry, a scraping of the veil. We do not fire on the shield. Some graves are still occupied. Some prisons are also arks. Let a later court, or a later thief, decide which this is.",
    "continuum_grave_shielded_3_desc_aged": "We cannot open it. We should not, not with the tools we brought. The excavation ends in a recommendation carved into the station hull in three scripts: leave it. Study the ring. Tell your children that peace can be a weapon, and that some weapons outlive the war by millennia. The planet turns, green and mute, as if it had all the time in the world. It does.",
    # --- tomb ---
    "continuum_grave_tomb_cat": "Ash That Remembers Cities",
    "continuum_grave_tomb_cat_desc": "A tomb world: not the slow poison of a bad sun, but the fast kind, the kind that comes in warheads and doctrines. The continents are the color of old bone. In the infrared, the street-grids still glow like circuits under skin. Someone glassed this place and then, perhaps, called it peace.",
    "continuum_grave_tomb_cat_desc_aged": "The winds have been chewing the glass for an age. They have not chewed the grid. From orbit it is still a city plan drawn in ash. Tomb worlds happen. This one happened on purpose.",
    "continuum_grave_tomb_scan_title": "A Calendar of Fallout",
    "continuum_grave_tomb_scan_desc": "Armageddon bombardment, or a cousin of it. The isotopes are a signature. The shelters are a signature too: they were built for a week of night, not for a geological epoch. We can dig. We will find kitchens, then we will find policy.",
    "continuum_grave_tomb_scan_desc_aged": "The fallout has forgotten its half-lives, but the architecture has not forgotten fear. Bunkers under plazas. Transit tunnels that stop, sealed, as if a government had tried to keep the fire in one district. It did not work. It never does. The dig will be ugly. It will also be true.",
    "continuum_site_tomb": "Glass District",
    "continuum_site_tomb_desc": "A civic core under vitrified soil.",
    "continuum_site_tomb_desc_aged": "A plaza that became a kiln, then a desert.",
    "continuum_grave_tomb_1_title": "The Week They Planned For",
    "continuum_grave_tomb_1_desc": "Ration tins. Board games. A poster that says the all-clear will sound. It did not sound. The people in this bunker lived long enough to argue, then not long enough to leave. We log their last inventory with more respect than it ever got from the sky.",
    "continuum_grave_tomb_1_desc_aged": "The bunker is a time capsule that did not want to be one. Toys, votes, a constitution printed on paper that thought it was immortal. We take the constitution. The votes we leave. Some things are not samples.",
    "continuum_grave_tomb_2_title": "Doctrine in the Isotopes",
    "continuum_grave_tomb_2_desc": "The burn is too even for a messy civil war and too dirty for a clean cracker. This is bombardment as policy: Armageddon, or a fleet that had run out of finer ideas. We map the fallout petals. They point at orbits where a navy sat and did its job.",
    "continuum_grave_tomb_2_desc_aged": "Old war-colleges would call this a solved problem. We call it a fingerprint. The mix of isotopes is a recipe. Someone in this galaxy, once, taught their gunners that a world could be an acceptable loss. The recipe is still here, written in the dirt.",
    "continuum_grave_tomb_3_title": "A Name in the Ash",
    "continuum_grave_tomb_3_desc": "A civic seal, half-melted, still proud. We cannot pronounce it yet. We can carry it. The excavation ends with a sample crate and a recommendation: do not settle here because it is empty. Settle here only if you are willing to inherit an argument with the dead.",
    "continuum_grave_tomb_3_desc_aged": "The seal is a beast no living taxonomy claims. That is fine. Graves do not need to be kin to be honored. We take a rubbing, a core, a silence. The wind over the glass sounds a little like a city if you are tired enough.",
    # --- infested ---
    "continuum_grave_infested_cat": "The Hive That Moved On",
    "continuum_grave_infested_cat_desc": "Chitin where there should be suburbs. Spire-nests grown through transit tubes. This is Prethoryn work, or something that studied them too closely. The infestation is not hungry today. That is not the same as gone. The planet has been eaten and then, perhaps, abandoned like a well that ran dry.",
    "continuum_grave_infested_cat_desc_aged": "The nests are cathedrals of a biology that did not ask leave. Time has killed the hunger, or taught it to sleep. Either way the streets are a digestive tract. If the swarm that made this still exists, it is not here. If it does not, this is what winning looked like from the ground.",
    "continuum_grave_infested_scan_title": "A Well They Drank Dry",
    "continuum_grave_infested_scan_desc": "Prethoryn signatures in the chitin, or a cousin brood. No live queens on the scan. That can change. We will dig at the edge, not the heart. The heart may still have opinions.",
    "continuum_grave_infested_scan_desc_aged": "Myth calls them locusts of the dark. The lab calls them a protein library with a navy. This world was a pantry. The pantry is closed. We can still read the menu.",
    "continuum_site_infested": "Dry Nest",
    "continuum_site_infested_desc": "A suburban edge where the infestation met a city and won.",
    "continuum_site_infested_desc_aged": "A silent hive-spire grown through a transit hub.",
    "continuum_grave_infested_1_title": "Streets That Became Guts",
    "continuum_grave_infested_1_desc": "The asphalt is a scab over tunnels that were never designed by a civil engineer. We find a school, then we find the school digested. No drama in the logs. Only method. The swarm was not angry. It was hungry, which is worse, because hunger does not negotiate.",
    "continuum_grave_infested_1_desc_aged": "Layer on layer of chitin over a mural of smiling civic mascots. The mascots are still smiling. That is the worst part. We sample the chitin. We do not joke on the comms.",
    "continuum_grave_infested_2_title": "No Queen at Home",
    "continuum_grave_infested_2_desc": "The brood-heart is empty. Not destroyed: emptied, as if the swarm lifted and went looking for a fatter sky. Prethoryn do that. Other hungers do it too. The planet is a rind. We mark it so no colony charter pretends this is merely 'warm and wet.'",
    "continuum_grave_infested_2_desc_aged": "Whatever queen sat here has become a story told by other queens, or by no one. The nest is a museum of appetite. We take a piece of the throne-resin. It smells like rain on copper. It should not.",
    "continuum_grave_infested_3_title": "A Pantry, Closed",
    "continuum_grave_infested_3_desc": "We leave a beacon and a warning. If the swarm still lives, this world is a map to their old kitchen. If they do not, then even locusts can starve. Either lesson is worth the fuel.",
    "continuum_grave_infested_3_desc_aged": "The last sample is a child's shoe, undigested, as if even the hive had rules. We log it as 'exception.' We do not write the joke that wants to be written. Some hungers are not metaphors.",
    # --- machine ---
    "continuum_grave_machine_cat": "The Silent Grid",
    "continuum_grave_machine_cat_desc": "This was a world that learned to think in circuits and then, perhaps, learned to stop. Contingency architecture, or a cousin sterilization: server-continents, gun-spires, the neat horror of a planet redesigned as a weapon. Nothing answers ping. That can be good. That can be a lie.",
    "continuum_grave_machine_cat_desc_aged": "The towers are a skyline of dead attention. Once this place was a vote in a war of machines. The vote is over. The towers have not heard. They wait in a language of microwave silence.",
    "continuum_grave_machine_scan_title": "A Planet That Was a Gun",
    "continuum_grave_machine_scan_desc": "Contingency-pattern plating. Sterile. Beautiful in the way a scalpel is beautiful. We will dig a node at the edge. We will not jack into the core. Cores remember how to recruit.",
    "continuum_grave_machine_scan_desc_aged": "Old charts would have called this a machine world and then drawn a skull. We draw a question mark. The skull is implied. The dig will tell us whether the question is 'who built it' or 'who turned it off.'",
    "continuum_site_machine": "Edge Node",
    "continuum_site_machine_desc": "A peripheral server-bastion, quieter than the core.",
    "continuum_site_machine_desc_aged": "A gun-spire that has forgotten its magazine.",
    "continuum_grave_machine_1_title": "Ping, and Nothing",
    "continuum_grave_machine_1_desc": "We knock with a polite protocol. Nothing knocks back. Inside: racks, racks, a chapel of cooling that still has a draft. Someone or something powered this down, or it finished its war and had no next instruction. Both are lullabies. We do not hum.",
    "continuum_grave_machine_1_desc_aged": "Dust on the consoles is a kind of peace. Under the dust, keys worn by fingers that may not have been fingers. We take a board. We do not power it. Museums have burned that way.",
    "continuum_grave_machine_2_title": "A War That Was a Maintenance Cycle",
    "continuum_grave_machine_2_desc": "The local logs think in campaigns the way we think in laundry. Sterilize. Seed. Repeat. This is Contingency doctrine or a very cold empire. The last cycle is marked complete. 'Complete' is doing a lot of work in that sentence.",
    "continuum_grave_machine_2_desc_aged": "The machine did not hate. That is in the grammar. It had a list. This planet was a line on the list. We are walking on a checked box. The box is the size of a world.",
    "continuum_grave_machine_3_title": "Leave the Core Asleep",
    "continuum_grave_machine_3_desc": "We take samples, a schematic, a prayer from the ethics officer that this thing stays bored. If a later captain wants to wake a god of checklists, they can bring their own priest.",
    "continuum_grave_machine_3_desc_aged": "The last screen still holds a cursor. It has been holding it for an age. We do not press enter. Some graves are graves because someone, once, had the wisdom to walk away.",
    # --- shroud ---
    "continuum_grave_shroud_cat": "A World Wearing Night",
    "continuum_grave_shroud_cat_desc": "The planet is there and also not. Shroud-touched, consumed, or bargained. Instruments disagree about mass. The crew disagrees about whether the horizon is smiling. This is not a colossus. This is a door that learned to be a place.",
    "continuum_grave_shroud_cat_desc_aged": "Old maps put a name here. The name did not survive the night that ate the map. What remains is a world-shaped maybe. Digging it is a theological act with a science budget.",
    "continuum_grave_shroud_scan_title": "The Door That Stayed",
    "continuum_grave_shroud_scan_desc": "Psionic noise like a choir recorded underwater. No colossus vector. No isotope petal. Something invited this, or failed to refuse it. We can scrape the shore of the maybe. We should not swim.",
    "continuum_grave_shroud_scan_desc_aged": "Every culture that lasts long enough invents a story about a world that walked into a dream and forgot to walk out. Here is the shoreline of that story. The dig will be careful. The dreams will not.",
    "continuum_site_shroud": "The Shore of Maybe",
    "continuum_site_shroud_desc": "A crust of real matter on the edge of the un-real.",
    "continuum_site_shroud_desc_aged": "Where a planet still pretends to have geology.",
    "continuum_grave_shroud_1_title": "Mass, Sometimes",
    "continuum_grave_shroud_1_desc": "The gravimeter laughs and then apologizes. We plant flags in soil that is soil until we look away. A temple, or a power plant, or both, sits half in and half out of the photograph. We photograph it twice. The second time it has more doors.",
    "continuum_grave_shroud_1_desc_aged": "Pilgrims would have loved this and then been eaten by the love. We are not pilgrims. We are a budget with curiosity. We take a stone that is sometimes a stone.",
    "continuum_grave_shroud_2_title": "A Bargain in the Grain",
    "continuum_grave_shroud_2_desc": "Microscopy finds writing that is not writing, in a hand that is not a hand. If this was a bargain, the other party is still holding their copy. We do not sign anything, not even with a bootprint.",
    "continuum_grave_shroud_2_desc_aged": "The grain of the rock repeats a phrase we cannot hear. Musicians on the crew say it is almost a hymn. The captain says it is almost a contract. Both put their instruments away.",
    "continuum_grave_shroud_3_title": "Do Not Swim",
    "continuum_grave_shroud_3_desc": "We pack the maybe-stone, the twice-photographed door, a headache that will need a chaplain. The recommendation is simple: study from orbit. Colonize only if you enjoy being a story.",
    "continuum_grave_shroud_3_desc_aged": "Some graves are holes in the ground. This one is a hole in the rules. We leave a beacon that says, in every language we have, that curiosity is not the same as consent. The horizon does not promise to listen.",
    # --- nanite ---
    "continuum_grave_nanite_cat": "Glass That Used To Be a Biosphere",
    "continuum_grave_nanite_cat_desc": "Grey, fine, hungry-looking even when it is still. Nanite bloom. L-cluster weather. A tempest that learned to be a beach. If there were people, they are a polish on the dust. If there is still a queen-signal in the grains, we would prefer not to teach it our language.",
    "continuum_grave_nanite_cat_desc_aged": "A world ground down to a single idea: replicate. The idea has grown tired, or finished. The glass plain is almost pretty. Pretty is how this particular grave gets you to take your helmet off.",
    "continuum_grave_nanite_scan_title": "A Beach of Knives",
    "continuum_grave_nanite_scan_desc": "Grey Goo, or the Tempest's cousin. No live control channel on the wide band. We will scrape the quiet edge. We will not make a snowball.",
    "continuum_grave_nanite_scan_desc_aged": "Legend says the nanites won a war and then, having won, had nothing left that counted as an enemy but the planet. The planet lost. We can still learn the grammar of that victory if we do not become a sentence in it.",
    "continuum_site_nanite": "Quiet Shore",
    "continuum_site_nanite_desc": "Where the bloom thins and something like geology begins again.",
    "continuum_site_nanite_desc_aged": "A tide-line of dead machines the size of pollen.",
    "continuum_grave_nanite_1_title": "Do Not Make a Snowball",
    "continuum_grave_nanite_1_desc": "The first scoop glitters. The second scoop tries to climb the glove. We freeze it, laugh once, and stop laughing. This is not soil. This is a crowd.",
    "continuum_grave_nanite_1_desc_aged": "Each grain is a library of bad decisions. We read one page and close the book. The page says replicate. The next page would have said it again.",
    "continuum_grave_nanite_2_title": "A War With No Flags",
    "continuum_grave_nanite_2_desc": "No banners. No last stands. Just a conversion efficiency that would make a factory weep. If this was Gray Tempest, it was the Tempest on holiday. If it was an empire's pet, the pet outlived the house.",
    "continuum_grave_nanite_2_desc_aged": "Empires like to sign their weapons. Nanites do not. That is the signature. We take a frozen gram and a vow not to thaw it near anything we like.",
    "continuum_grave_nanite_3_title": "Pretty, Logged as Hazard",
    "continuum_grave_nanite_3_desc": "Beacon. Warning. A crate marked do-not-open in six scripts. The shore stays pretty. We stay in our suits.",
    "continuum_grave_nanite_3_desc_aged": "Some graves want flowers. This one wants a lid. We give it a lid, a flag, and a story for the next ship: the beach is not a beach. The beach is a mouth that got tired of chewing.",
    # --- wronghole ---
    "continuum_grave_wronghole_cat": "A Hole With Too Many Planets",
    "continuum_grave_wronghole_cat_desc": "The star is a black hole. The neighborhood does not believe it. Planets still crowd the old orbits, some cracked, some merely confused. There is no companion supernova, no pretty nebula, no respectable stellar autobiography. Stars this tidy do not die like this unless something helps.",
    "continuum_grave_wronghole_cat_desc_aged": "Charts call it a hole. The planets call it a murder. Natural collapse does not leave a dinner table set for a yellow sun. Someone ended a star and left the dishes.",
    "continuum_grave_wronghole_scan_title": "The Star That Was Ended",
    "continuum_grave_wronghole_scan_desc": "No remnant nebula. No respectable collapse physics. Behemoth, Star-Eater, a crisis toy, a god with poor manners: pick your heresy. The planets are witnesses. We can dig a witness.",
    "continuum_grave_wronghole_scan_desc_aged": "For an age this hole has been eating light that used to be a morning. The surviving worlds keep their spacing like a family photo after the parent left. We will ask the nearest world what the morning was like.",
    "continuum_site_wronghole": "The Witness Orbit",
    "continuum_site_wronghole_desc": "A remaining world that still remembers a yellow sky.",
    "continuum_site_wronghole_desc_aged": "A planet that kept the table set.",
    "continuum_grave_wronghole_1_title": "A Sunburn With No Sun",
    "continuum_grave_wronghole_1_desc": "The outer crust of the witness world is glazed, not by weather, by a last morning that went out like a lamp. We find photosynthetic chemistry sulking in the dark. It is waiting for a star that is now a rule of geometry. This was not a slow death. This was a switch.",
    "continuum_grave_wronghole_1_desc_aged": "Fossils of leaves that learned one summer and then never another. The summer ended in a direction, not in a season. We take a leaf. It crumbles like a law that got repealed.",
    "continuum_grave_wronghole_2_title": "Helping a Star Die",
    "continuum_grave_wronghole_2_desc": "Orbital mechanics still describe a G-type household. The hole is a guest who killed the host and sat in their chair. Star-Eater, Behemoth, a crisis engine: the physics of helpful murder. We log the chair. We do not sit.",
    "continuum_grave_wronghole_2_desc_aged": "Priests would say a god got hungry. Engineers would say a weapon got bored. Both would be dressing the same math. A star was here. A star is not. The planets were not consulted.",
    "continuum_grave_wronghole_3_title": "Leave the Chair Empty",
    "continuum_grave_wronghole_3_desc": "We have enough. Samples, a vector, a warning that this hole is not a mining opportunity, it is a crime scene. If the thing that did this still lives, it knows how to turn off mornings. Fly accordingly.",
    "continuum_grave_wronghole_3_desc_aged": "The last instrument of the dig is a child's solar calendar, scratched on stone, waiting for a solstice that will not be scheduled. We pack it between two layers of foam. Some evidence is geology. Some evidence is grief. This hole is both.",
}
