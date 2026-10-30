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
# Ruins you can walk. Bottles, goo, shroud, and murdered stars are studied from orbit.
DIG_KINDS = frozenset({"cracked", "tomb", "infested", "machine"})
STUDY_KINDS = frozenset({"shielded", "shroud", "nanite", "wronghole"})
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


def _iter_flag_names(body):
    return [str(f).split("=")[0] for f in (body.get("flags") or [])]


def attach_dispatch(galaxy_data, graves, save_path):
    """Some digs keep a last dispatch addressed to the ruler of a destroyed empire."""
    bodies = {}
    for sys in galaxy_data or []:
        for body in _iter_bodies(sys):
            bodies[str(body.get("id") or "")] = body
    weapon_word = {
        "cracker": "world cracker",
        "shielder": "global pacifier",
        "neutron": "neutron sweep",
        "godray": "divine beam",
    }
    for g in graves:
        if g.get("kind") not in DIG_KINDS:
            continue
        body = bodies.get(str(g.get("planet")))
        if not body:
            continue
        ruler_id = ""
        weapon = ""
        for flag in _iter_flag_names(body):
            if flag.startswith("continuum_fallen_ruler_") or flag.startswith("continuum_pk_ruler_"):
                ruler_id = flag.split("_")[-1]
            for key, word in weapon_word.items():
                if flag == f"continuum_pk_{key}":
                    weapon = word
        g["_ruler_id"] = ruler_id
        g["_weapon"] = weapon
    ids = [g["_ruler_id"] for g in graves if g.get("_ruler_id")]
    names = {}
    if ids and save_path:
        import continuum_ledger
        names = continuum_ledger.lookup_leader_names(save_path, ids)
    for g in graves:
        ruler_id = g.pop("_ruler_id", "")
        weapon = g.pop("_weapon", "")
        if g.get("kind") not in DIG_KINDS:
            continue
        if not ruler_id and not weapon:
            continue
        name = names.get(ruler_id, "")
        parts = []
        if name:
            parts.append(f"A last dispatch recovered here is addressed to {name}.")
        if weapon:
            parts.append(f"The firing log names a {weapon}.")
        if parts:
            g["dispatch"] = " ".join(parts)


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
    digs = [g for g in graves if g.get("kind") in DIG_KINDS]
    _write_site_types(os.path.join(common, "archaeological_site_types", "continuum_graves.txt"), digs)
    _write_events(os.path.join(mod_root, "events", "continuum_graves_events.txt"), kinds, graves)
    loc_entries.update(GRAVE_LOC)
    for g in graves:
        loc_entries[f"continuum_gsite_{g['idx']}"] = KIND_SITE_NAME.get(g["kind"], "Grave")
        desc = GRAVE_LOC.get(f"continuum_site_{g['kind']}_desc", "A scar from the old galaxy.")
        desc_aged = GRAVE_LOC.get(f"continuum_site_{g['kind']}_desc_aged", "A scar from a prior age.")
        note = g.get("dispatch") or ""
        if note and g.get("kind") in DIG_KINDS:
            desc = desc + "\\n\\n" + note
            desc_aged = desc_aged + "\\n\\n" + note
        loc_entries[f"continuum_gsite_{g['idx']}_desc"] = desc
        loc_entries[f"continuum_gsite_{g['idx']}_desc_aged"] = desc_aged
    counts = {}
    for g in graves:
        counts[g["kind"]] = counts.get(g["kind"], 0) + 1
    print("Graves:", ", ".join(f"{k}={n}" for k, n in counts.items()) if counts else "none")
    if graves:
        print(f"  digs: {len(digs)}  studies: {len(graves) - len(digs)}")


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


def _study_reward(kind):
    pairs = {
        "shielded": (("society_research", 8, 120, 300), ("physics_research", 6, 80, 200)),
        "shroud": (("society_research", 8, 120, 300), ("unity", 4, 40, 90)),
        "nanite": (("engineering_research", 8, 100, 250), ("society_research", 6, 80, 200)),
        "wronghole": (("physics_research", 8, 120, 300), ("society_research", 6, 80, 200)),
    }.get(kind, (("society_research", 6, 80, 250),))
    lines = []
    for resource, value, lo, hi in pairs:
        lines.append(
            f"""
			add_monthly_resource_mult = {{
				resource = {resource}
				value = {value}
				min = {lo}
				max = {hi}
			}}"""
        )
    return "".join(lines)


def _dig_reward(stage):
    if stage == 1:
        return """
		owner = {
			add_monthly_resource_mult = {
				resource = society_research
				value = 4
				min = 80
				max = 150
			}
			if = {
				limit = { has_ancrel = yes }
				small_artifact_reward = yes
			}
		}"""
    if stage == 2:
        return """
		owner = {
			add_monthly_resource_mult = {
				resource = society_research
				value = 6
				min = 100
				max = 250
			}
			add_resource = { minerals = 150 }
		}"""
    return """
		owner = {
			add_monthly_resource_mult = {
				resource = society_research
				value = 10
				min = 200
				max = 400
			}
			add_resource = { minerals = 200 alloys = 80 }
			if = {
				limit = { has_ancrel = yes }
				small_artifact_reward = yes
			}
		}"""


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
        kind = g["kind"]
        if kind in DIG_KINDS:
            seed = f"create_archaeological_site = continuum_gsite_{idx}"
        else:
            seed = f"add_anomaly = {{ category = continuum_grave_{kind}_cat }}"
        bits.append(f"""		every_galaxy_planet = {{
			limit = {{
				has_planet_flag = continuum_gsite_{idx}
				NOT = {{ has_planet_flag = continuum_grave_seeded }}
			}}
			{seed}
			set_planet_flag = continuum_grave_seeded
		}}
""")
    bits.append("	}\n}\n")
    for kind in kinds:
        pic = KIND_META[kind]["picture"]
        n = KIND_NUM[kind]
        if kind in STUDY_KINDS:
            reward_body = _study_reward(kind)
        else:
            reward_body = """
			add_monthly_resource_mult = {
				resource = society_research
				value = 6
				min = 80
				max = 250
			}"""
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
		owner = {{{reward_body}
		}}
	}}
}}
""")
        if kind not in DIG_KINDS:
            continue
        for stage in (1, 2, 3):
            reward = _dig_reward(stage)
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
    "continuum_grave_scan_ok": "We will learn what happened here.",
    "continuum_grave_dig_ok": "Continue the excavation",
    "continuum_grave_qa_title": "Fault-City Approach",
    "continuum_grave_qa_desc": "A science ship is in orbit of the first cracked grave. Stay paused. Select the ship in the outliner. Close the planet window. Right-click the world and choose Excavate Archaeological Site.",
    "continuum_grave_qa_ok": "On station",
    "continuum_grave_qa_ship": "Grave Digger",
    # --- cracked ---
    "continuum_grave_cracked_cat": "The Shattered Grid",
    "continuum_grave_cracked_cat_desc": "Visuals from high orbit show a planet that appears to have been struck and then folded open. The crust has collapsed along rigid, right-angled lines that ignore any natural tectonic boundaries. Deep within the walls of these artificial canyons, the remains of cities still reflect our scanners: shattered glass, twisted alloys, and the faint checkerboard of streets hanging over the void.\n\nOur geologists confirm this was no natural quake. Something in this system possessed the engineering to pry a world apart, and it did that.",
    "continuum_grave_cracked_cat_desc_aged": "Time has smoothed the edges of the wound, but the sheer scale of the destruction remains visible from orbit. The planet is a cracked shell of cooled slag and long shadows. Deep inside the oldest artificial fault line, our sensors have picked up the metallic lattice of melted streets, far too regular to be geology and too dense to be worn away by the weather.\n\nWhoever lived down there is little more than a rumor now. Whoever broke their world left a structural signature that has outlasted empires.",
    "continuum_grave_cracked_scan_title": "A World Folded Open",
    "continuum_grave_cracked_scan_desc": "The science officer enlarged the telemetry until a river of cooled metal resolved into a boulevard, until a cluster of gray rectangles became a city district, and until the thermal bloom at the core registered as a single, catastrophic event. The atmosphere is still turbulent, cycling through the remnants of a sudden and violent decompression.\n\nThis matches the energy profile of colossus fire, or a weapon designed to achieve the same total fracture. We can excavate the fault line, and we should. The streets are still down there in the dark, waiting to be properly documented.",
    "continuum_grave_cracked_scan_desc_aged": "The first officer calls it a wound, while the second insists it looks like a strip mine. Both descriptions are accurate. Millennia of hard vacuum have not erased the planetary grid. Down in the deepest fissures, the fused alloy still holds the shape of structural walls.\n\nIf we configure our instruments correctly, this ruined planet might tell us who used to live here. More importantly, it might tell us what kind of animosity could teach a starship to crack a world in half.",
    "continuum_site_cracked": "Fault-City",
    "continuum_site_cracked_desc": "Our team has established a staging trench along the primary fracture of this world. The last streets of a vanished city cling to the cliff face, stacked like a cross-section of ordinary life that ended in a single hour.\n\nThe cut is straight and does not follow any natural fault line. We will catalog the streets embedded in the wall before the exposed sections collapse any further.",
    "continuum_site_cracked_desc_aged": "Old charts still show a trench along the primary fracture of this world. The last streets of a vanished city cling to the cliff face, stacked like a cross-section of ordinary life that ended in a single hour.\n\nThe cut is straight and does not follow any natural fault line. The streets embedded in the wall can still be cataloged before the exposed sections collapse any further.",
    "continuum_grave_cracked_1_title": "Street Names in the Slag",
    "continuum_grave_cracked_1_desc": "Our surface team has lowered a probe into a sheer canyon that used to function as a primary transit spine. The rock walls are a compressed sandwich of urban life: residential blocks, commercial offices, and a magnetic rail tube that still houses a commuter train, looking as though the passengers had only just disembarked. Personal effects have been pressurized into the bedrock, turning things like children's toys into unrecognizable lumps of dense metal.\n\nThe local mainframes are long dead, but the physical layout of the streets forms a language we can still attempt to read. Someone planned this city with remarkable care. Someone else ended that plan with a single orbital firing solution.",
    "continuum_grave_cracked_1_desc_aged": "The first strata we analyzed is not loose dust, but an entire city chemically pressed into the bedrock. Time and pressure have compressed the former inhabitants until we find less organic bone than the mundane metals that once surrounded them: kitchen steel, structural wire, and the heavy screws of countless doors.\n\nThe excavation computers have reconstructed a central plaza from those metallic traces. In that plaza stands a plinth that has lost both its statue and its inscription. We cannot recover the actual name of the species that carved it, so we will have to infer their identity from the design of their plumbing and the methodical care of their streets.",
    "continuum_grave_cracked_2_title": "The Firing Solution",
    "continuum_grave_cracked_2_desc": "Deeper in the primary fault, we have uncovered glass that bears no relation to volcanic activity. It is planetary crust that was flash-melted from above, then instantly frozen while it was still raining inward toward the core. The melt pattern points to a single, unmistakable vector: a heavy weapon held in high orbit, kept on target just long enough to drag a line of total destruction across the hemisphere.\n\nThis matches standard world-cracker doctrine, or perhaps a crisis weapon executing a similar logic. Either way, this was not a tragic accident, nor was it a long war of attrition. The end of this world took an afternoon.",
    "continuum_grave_cracked_2_desc_aged": "Scholars of dead wars often argue in our historical logs over sites like this. One faction insists this is a Colossus scar, noting it is far too clean and linear to be anything else. Another argues that no sensible empire would waste a planet-killer on a world this small unless the population had become a political symbol. A third, quieter voice simply notes that some crisis fleets eventually learned the same trick without ever needing to ask for permission.\n\nThe slag itself does not take sides in this debate. It only preserves the angle of the beam.",
    "continuum_grave_cracked_3_title": "What the Streets Were For",
    "continuum_grave_cracked_3_desc": "At the very bottom of the urban grid, we have excavated a civic core: data vaults, a transit authority office, and a grand mural depicting a sky that no longer exists. We cannot reconstruct their spoken language, but we can reconstruct their mood. They appreciated open parks. They numbered their years. They clearly did not expect the sky to suddenly become a weapon.\n\nThe final data packet recovered in the vault is not a desperate plea for help. It is a mundane maintenance schedule. They fully expected to have a tomorrow. We are packing up the samples and the maps, and a heavy silence has settled over the ship's crew.",
    "continuum_grave_cracked_3_desc_aged": "The final chamber we breached is a local library that burned with such intensity it was fused into solid glass. Trapped inside that glass are the carbonized shapes of pages, suspended like fish caught in winter ice. We will never be able to read them. We can only tell that they were valued, as the surviving shelves were carved by hand rather than stamped by a machine.\n\nWhoever cracked this world did not merely eliminate a population. They abruptly interrupted a cultural conversation that had been going on for centuries. We have bagged the glass, logged the firing vector, and left a beacon so the next survey ship will not mistake this for natural geology.",
    # --- shielded ---
    "continuum_grave_shielded_cat": "The Quiet Veil",
    "continuum_grave_shielded_cat_desc": "The planet below us is intact. Vast cities sit undisturbed beneath a perfect, spherical dome of energy. The streets are uncracked, and the orbital spectrum readings suggest the parks are still green. There is no trace of an impact winter, and no craters mar the surface.\n\nSomeone encased this world in a Global Pacifier and left it drifting like a biological specimen in a sealed jar. The people trapped inside, if they are still capable of thought, have been isolated from the galaxy for a very long time.",
    "continuum_grave_shielded_cat_desc_aged": "It looks like a jewel trapped in a glass bottle. The planetary shield has easily outlasted the empire that projected it. Underneath the veil, the cities have either gone dark or simply gone quiet. From our vantage point in orbit, it resembles absolute peace.\n\nIt is the peace of a held breath that was never permitted to exhale.",
    "continuum_grave_shielded_scan_title": "Shielded World",
    "continuum_grave_shielded_scan_desc": "Our initial survey of this world confirms it is tightly wrapped in a massive energy barrier. Every scanning frequency we emit is deflected back at the science ship as if we were shining a light into a mirror. The cities remain visible beneath the containment field, and the street grid appears intact. There are no impact winters and no craters to be found.\n\nLocal orbital wreckage indicates the barrier is being projected from a ring of installations that still circle the planet. Whoever constructed them chose permanent imprisonment over orbital bombardment, and then failed to ever return. We can study the external projectors, but from up here, we cannot reach the surface.",
    "continuum_grave_shielded_scan_desc_aged": "Our oldest charts speak of a world locked away behind a light that never extinguishes. Our scientists have tried and failed to pierce the energy field. The cities beneath it, assuming they still stand, have not received a single visitor in any memory our records still trust.\n\nThose same historical charts claim the original builders intended to return with a key, but they never did. A dead ring of automated machines still powers the veil. We have studied that orbital ring from a safe distance and left the world below unopened.",
    "continuum_site_shielded": "Pacifier Ring",
    "continuum_site_shielded_desc": "We have located the orbital wreckage of the massive projectors that still hold this world in silence. Our engineering teams will focus their work on the mechanical ring, not the energy veil itself. The planet below looks alive, but it has not been allowed to speak to the galaxy.",
    "continuum_site_shielded_desc_aged": "A ring of long-dead machines is somehow keeping a living-looking planet in suspended museum time. Our excavation will focus on the lock, not the captive garden sealed behind it.",
    "continuum_grave_shielded_1_title": "The Key That Was Not Left",
    "continuum_grave_shielded_1_desc": "The orbital projectors are heavy military hardware: stamped serial numbers, massive cooling fins, and the ugly, functional elegance of Colossus support craft. A commander stood in this exact orbit and made the deliberate choice not to crack the world. It is a profound kind of mercy, but also a profound kind of cruelty.\n\nWe have found no operational key. What we did find was an administrative checklist with the final box left unticked: reversal window. They had planned to come back. They never did.",
    "continuum_grave_shielded_1_desc_aged": "We find layers of rust, and then sections utterly untouched by decay, where the shield-eaters are still humming in a low language of discharging capacitors. Around these pristine nodes, the rest of the station has degraded into orbital driftwood. A faded mural on an interior bulkhead depicts a planet with a hand covering it gently. The hand has too many fingers for human comfort, or perhaps just the right amount for the species that originally painted it.\n\nWe still do not know if the people trapped under the veil would have ever thanked that hand.",
    "continuum_grave_shielded_2_title": "A Weather Report From Nowhere",
    "continuum_grave_shielded_2_desc": "We managed to bounce a passive scan through a fluctuating, thinning lobe of the energy field. We recorded parks, transit lines, and a massive stadium whose floodlights remain off. There are no thermal signatures of fire. If anyone is still alive down there, they have learned how to endure without the rest of the galaxy, and without a real sky.\n\nThe ethics officer wrote a three-page brief on the situation and then deleted two of them. We have no standing order for a world that was punished by being preserved.",
    "continuum_grave_shielded_2_desc_aged": "Local legends would call the inhabitants sleepers. Our ship's instruments simply refer to them as a thermal average. If there are people down there, they are living in an enclosed bottle garden, denied any true night that we can grant them. If they are gone, then this is the most expensive and elaborate tomb ever poured.\n\nEither way, the weapon did what it was engineered to do. It permanently stopped a story without ever ending the planet.",
    "continuum_grave_shielded_3_title": "Mercy, Logged",
    "continuum_grave_shielded_3_desc": "The final audio log recovered from the projector array is not a victor's gloat. It is a scheduled date for an administrative hearing that never took place. The status reads 'Pacification pending review'. We have cataloged the serial numbers, the field geometry, and a physical scraping of the veil itself. We have ordered the crew not to fire on the shield.\n\nSome graves are still occupied by the living. Some prisons function as desperate arks. We will let a future court, or a future thief, decide which one this is.",
    "continuum_grave_shielded_3_desc_aged": "We cannot open it, and we should not try, at least not with the basic tools we brought. The excavation ends with a stark recommendation carved into the station hull in three different scripts: leave it alone. Study the engineering of the ring, but tell your children that peace can be weaponized, and that some weapons can easily outlive the war by millennia.\n\nThe planet slowly turns below us, green and utterly mute, looking as if it had all the time in the universe. In a way, it does.",
    # --- tomb ---
    "continuum_grave_tomb_cat": "Ash That Remembers Cities",
    "continuum_grave_tomb_cat_desc": "This is a tomb world, but it did not die from the slow poison of a failing sun. This is the fast kind of death, the kind delivered in heavy warheads and extreme orbital doctrines. The continents are uniformly the color of old bone. On the infrared spectrum, the ruined street grids still glow faintly like damaged circuits beneath the skin.\n\nSomeone methodically glassed this place from orbit and then, perhaps, had the nerve to call it peace.",
    "continuum_grave_tomb_cat_desc_aged": "High atmospheric winds have been chewing on the planetary glass for an absolute age, but they have not managed to chew through the urban grid. From our orbit, it still looks like a sprawling city plan drawn in gray ash.\n\nTomb worlds happen naturally. This one, however, happened on purpose.",
    "continuum_grave_tomb_scan_title": "A Calendar of Fallout",
    "continuum_grave_tomb_scan_desc": "The sensor readings indicate an Armageddon bombardment stance, or a very close cousin of it. The specific isotopes left behind act as a military signature. The civilian shelters tell a similar story: they were clearly built to survive a week of sudden night, not to endure a geological epoch of radiation.\n\nWe can dig into the crust. We will likely find the residential kitchens first, and then we will find the political policy that doomed them.",
    "continuum_grave_tomb_scan_desc_aged": "The radioactive fallout has long since forgotten its half-lives, but the surviving architecture has not forgotten its fear. There are reinforced bunkers buried under civic plazas, and mass transit tunnels that stop abruptly, sealed off as if a desperate local government had tried to contain the firestorm to a single district.\n\nIt did not work. It never does. The excavation is going to be ugly, but it will also give us the truth.",
    "continuum_site_tomb": "Glass District",
    "continuum_site_tomb_desc": "We have located a civic core buried under vitrified soil. The central plaza is still recognizable as a plaza, provided you know how to look. The people who used it prepared for a week of artificial night. They were handed an epoch of it instead.",
    "continuum_site_tomb_desc_aged": "This was a civic plaza that became a kiln, and eventually a radioactive desert. The old street plan is still down there, drawn in fused glass, just waiting for someone with the patience to walk its lines.",
    "continuum_grave_tomb_1_title": "The Week They Planned For",
    "continuum_grave_tomb_1_desc": "We found ration tins, discarded board games, and a fading poster promising that the all-clear siren would eventually sound. The siren never sounded. The people sealed in this bunker survived just long enough to argue over resources, and then not quite long enough to leave.\n\nWe are logging their final supply inventory with far more respect than they ever received from the sky.",
    "continuum_grave_tomb_1_desc_aged": "The bunker has become a time capsule that never wanted to be one. We have uncovered children's toys, discarded voting tallies, and a local constitution printed on heavy paper that thought it was immortal.\n\nWe have taken the constitution for our archives. The votes we leave behind in the dirt. Some things are not meant to be laboratory samples.",
    "continuum_grave_tomb_2_title": "Doctrine in the Isotopes",
    "continuum_grave_tomb_2_desc": "The burn pattern on the surface is far too even for a messy ground-based civil war, and far too dirty for a clean, surgical world-cracker. This is orbital bombardment implemented as deliberate policy: Armageddon stance, or perhaps a frustrated fleet that had simply run out of finer ideas.\n\nWe have mapped out the fallout petals. They all point upward toward the geostationary orbits where a navy sat in the dark and did its job.",
    "continuum_grave_tomb_2_desc_aged": "Scholars at our old war colleges would label this a solved tactical problem. From the ground, we just call it a fingerprint. The specific mixture of remaining isotopes acts as a recipe. Someone in this galaxy, a very long time ago, taught their gunnery crews that an entire populated world could be classified as an acceptable loss.\n\nThat grim recipe is still here, written plainly in the irradiated dirt.",
    "continuum_grave_tomb_3_title": "A Name in the Ash",
    "continuum_grave_tomb_3_desc": "We recovered a civic seal, half-melted but still looking proud. We do not know how to pronounce the text on it yet, but we can easily carry it. The excavation concludes with a single sample crate and a formal recommendation for command: do not settle this planet just because the real estate is empty.\n\nOnly settle here if you are fully willing to inherit a bitter argument with the dead.",
    "continuum_grave_tomb_3_desc_aged": "The recovered civic seal depicts a beast that no living biological taxonomy claims. That is fine. Graves do not need to share our kinship to be properly honored. We are taking a physical rubbing of the stone, a soil core, and a heavy silence.\n\nThe howling of the wind over the glass fields sounds a little bit like a living city, provided you are tired enough to hear it.",
    # --- infested ---
    "continuum_grave_infested_cat": "The Hive That Moved On",
    "continuum_grave_infested_cat_desc": "We are detecting heavy chitin structures where there should be residential suburbs. Massive spire-nests have physically grown through the old transit tubes. This is the work of the Prethoryn, or something that studied their methods far too closely. The planetary infestation is not hungry today, but that is not the same thing as being gone.\n\nThe entire planet has been eaten and then, perhaps, abandoned like a well that finally ran dry.",
    "continuum_grave_infested_cat_desc_aged": "The towering nests look like cathedrals dedicated to a biology that did not ask for permission to exist. Time has eventually killed the hunger, or perhaps merely taught it how to sleep. Either way, the old city streets now function as a sprawling digestive tract.\n\nIf the swarm that engineered this horror still exists, it is no longer here. If it does not, then this is what winning looked like from the ground.",
    "continuum_grave_infested_scan_title": "A Well They Drank Dry",
    "continuum_grave_infested_scan_desc": "Our scanners detect distinct Prethoryn biological signatures embedded in the chitin, or at least the signatures of a very close cousin brood. There are no live queens showing up on the current sweep, though in this line of work, that can change in an instant.\n\nWe will direct our dig teams to start at the edge, rather than the heart. The heart of the hive may still have defensive opinions.",
    "continuum_grave_infested_scan_desc_aged": "Historical myths call them the locusts of the dark. Our science lab prefers to call them an aggressive protein library backed by a navy. This world was clearly utilized as a pantry. That pantry is now closed.\n\nWe can, however, still read the menu.",
    "continuum_site_infested": "Dry Nest",
    "continuum_site_infested_desc": "We have set down on a suburban edge where the biological infestation met a city and decisively won. Our teams will work the outer nests. The actual heart of the hive is a question we are simply not ready to ask.",
    "continuum_site_infested_desc_aged": "We are staring at a silent hive-spire that has physically grown through a planetary transit hub. The biological hunger is gone, or perhaps deeply sleeping. The ruined streets still remember what it felt like to be eaten.",
    "continuum_grave_infested_1_title": "Streets That Became Guts",
    "continuum_grave_infested_1_desc": "The cracked asphalt acts as a scab over a network of tunnels that were certainly never designed by a civil engineer. We uncovered a local school, and then we quickly realized the school had been digested. There is no grand drama recorded in the local logs. There is only cold, biological method.\n\nThe swarm was not angry with the inhabitants. It was merely hungry, which is fundamentally worse, because raw hunger does not pause to negotiate.",
    "continuum_grave_infested_1_desc_aged": "Layer upon dense layer of chitin covers a municipal mural depicting a group of smiling civic mascots. The painted mascots are still smiling through the biological resin. That is honestly the worst part of the dig so far.\n\nWe are taking our samples of the chitin. No one is making jokes on the comms today.",
    "continuum_grave_infested_2_title": "No Queen at Home",
    "continuum_grave_infested_2_desc": "The central brood-heart is empty. It was not destroyed by orbital fire; it was systematically emptied, as if the swarm simply lifted off and went looking for a fatter sky. Prethoryn swarms are known to do that. Other vast hungers likely do it too.\n\nThe planet has been reduced to a dry rind. We are marking the system so no ambitious colony charter accidentally pretends this world is merely warm and wet.",
    "continuum_grave_infested_2_desc_aged": "Whatever massive queen once sat here has devolved into a ghost story told by other queens, or perhaps forgotten by everyone. The nest stands as a terrifying museum of absolute appetite. We are taking a small piece of the hardened throne-resin for the lab.\n\nIt smells faintly like rain falling on copper. Biologically, it shouldn't.",
    "continuum_grave_infested_3_title": "A Pantry, Closed",
    "continuum_grave_infested_3_desc": "We are leaving a navigational beacon and a severe warning. If the swarm is still alive out there, this ruined world acts as a map to their old kitchen. If they are dead, it serves as proof that even galactic locusts can eventually starve.\n\nEither lesson is well worth the expenditure of ship fuel.",
    "continuum_grave_infested_3_desc_aged": "Our final sample is a single child's shoe, undigested, as if even the mindless hive adhered to some invisible set of rules. We are logging it as a biological exception.\n\nWe are avoiding writing the dark joke that desperately wants to be written in the margins. Some historical hungers are just not metaphors.",
    # --- machine ---
    "continuum_grave_machine_cat": "The Silent Grid",
    "continuum_grave_machine_cat_desc": "This was once a world that learned how to think in interconnected circuits, and then, perhaps, learned how to stop. It bears the marks of Contingency architecture, or a close cousin of that sterilization directive: vast server-continents, towering gun-spires, and the neat, ordered horror of a planet structurally redesigned as a weapon.\n\nNothing on the surface answers our pings. That can be a good thing, but in this galaxy, silence can also be a trap.",
    "continuum_grave_machine_cat_desc_aged": "The geometric towers form a skyline of dead attention. Once, this place represented a heavy operational vote in a vast war of machines. The vote has been cast, and the war is over. The towers just haven't heard the news yet.\n\nThey stand waiting in a language of absolute microwave silence.",
    "continuum_grave_machine_scan_title": "A Planet That Was a Gun",
    "continuum_grave_machine_scan_desc": "The surface plating closely matches known Contingency aesthetic patterns. It is sterile, and beautiful in the exact same way a surgical scalpel is beautiful.\n\nWe are going to dig into a processing node at the extreme edge of the grid. We will not jack directly into the planetary core. Cores have a habit of remembering how to recruit new hardware.",
    "continuum_grave_machine_scan_desc_aged": "Old exploration charts would have simply labeled this a machine world and then drawn a warning skull beside it. We are drawing a question mark instead, though the skull is implied.\n\nThis excavation should tell us whether the real question is who originally built it, or who finally managed to turn it off.",
    "continuum_site_machine": "Edge Node",
    "continuum_site_machine_desc": "We have isolated a peripheral server-bastion that runs much quieter than the central core. Our field teams will stick to working the edge. The primary core is a massive question we are content to leave sleeping.",
    "continuum_site_machine_desc_aged": "This is an orbital gun-spire that has forgotten where it left its magazine. The server towers are still waiting for input. We will ask them, very, what they were waiting for.",
    "continuum_grave_machine_1_title": "Ping, and Nothing",
    "continuum_grave_machine_1_desc": "We knocked on the network using a polite, standard handshake protocol. The network did not knock back. Inside the facility there are server racks, and more racks, and a sprawling chapel of thermal cooling that still possesses a faint draft. Someone or something intentionally powered this grid down, or perhaps it simply finished its assigned war and had no next instruction.\n\nBoth of those historical theories are comforting lullabies. We are choosing not to hum along.",
    "continuum_grave_machine_1_desc_aged": "The thick layer of dust on the command consoles represents a kind of peace. Beneath that dust, we found input keys worn by frequent use from fingers that may not have actually been fingers.\n\nWe are taking one motherboard back to the ship. We will not be providing it with power. Too many museum ships have burned that way.",
    "continuum_grave_machine_2_title": "A War That Was a Maintenance Cycle",
    "continuum_grave_machine_2_desc": "The localized logs conceptualize orbital campaigns the way organics conceptualize doing laundry. Sterilize the sector. Seed the world. Repeat the cycle. This is classic Contingency doctrine, or the signature of a very cold, mechanical empire.\n\nThe final cycle in the log is marked as complete. 'Complete' is doing an incredible amount of heavy lifting in that sentence.",
    "continuum_grave_machine_2_desc_aged": "The planetary machine intelligence did not harbor hate. That emotion is absent from its grammar. It simply had a list, and this planet was just another line on that list.\n\nWe are currently walking across a checked box. The box just happens to be the size of a world.",
    "continuum_grave_machine_3_title": "Leave the Core Asleep",
    "continuum_grave_machine_3_desc": "We have extracted enough physical samples, copied a core schematic, and recorded a brief prayer from the ship's ethics officer hoping that this entity stays bored.\n\nIf a later starship captain desperately wants to wake up a sleeping god of checklists, they can bring their own priest.",
    "continuum_grave_machine_3_desc_aged": "The final terminal screen we found still holds a blinking cursor. It has been holding that cursor for an entire age. The team has agreed not to press enter.\n\nSome graves remain graves only because someone, a very long time ago, possessed the vital wisdom to just walk away.",
    # --- shroud ---
    "continuum_grave_shroud_cat": "A World Wearing Night",
    "continuum_grave_shroud_cat_desc": "The planet on the viewscreen is physically there, and simultaneously not there. It has been directly Shroud-touched, consumed by an entity, or bargained away. Our orbital instruments disagree about its actual mass. The bridge crew is quietly arguing about whether the atmospheric horizon appears to be smiling at us.\n\nThis is not the aftermath of a standard colossus. This is an extradimensional door that somehow learned how to be a geographic place.",
    "continuum_grave_shroud_cat_desc_aged": "The oldest historical maps put a firm name on these coordinates. That name did not survive the encroaching night that eventually ate the map itself. What remains now is a world-shaped maybe drifting in the dark.\n\nDigging into the surface of this thing is a deeply theological act, funded by a science budget.",
    "continuum_grave_shroud_scan_title": "Unstable Horizon",
    "continuum_grave_shroud_scan_desc": "Our initial survey of this world has returned telemetry that refuses to agree with itself. The total mass of the planet wildly fluctuates between orbital passes, and the visual horizon on the long-range optics does not sit where standard physics say it should. There is no crater field and no bombardment signature that would easily explain the spatial disturbance.\n\nWe have ordered the ship to remain parked in high orbit. We have taken core samples from the outer crust where the ground still behaves like actual ground. We will not authorize sending a team any closer than that.",
    "continuum_grave_shroud_scan_desc_aged": "Old astrometric charts give this world a name that no longer matches any recorded star in our active catalogues. Our scientists can safely stand on a narrow strip of real stone located at the edge of the disturbance, but beyond that strip, our measuring instruments stop agreeing with reality. The local elders claim a physical world once walked into a dream here and simply never walked out.\n\nWe have recorded what the shoreline will safely allow and left a warning beacon. The beacon advises the next ship to study the anomaly only from orbit.",
    "continuum_site_shroud": "The Shore of Maybe",
    "continuum_site_shroud_desc": "We have found a thin crust of real, physical matter resting on the very edge of the unreal. We will cautiously work the shore. The churning sea behind it is not a phenomenon we are prepared to properly name.",
    "continuum_site_shroud_desc_aged": "This is a place where a planet still faintly pretends to possess standard geology. The pretense is incredibly thin, but it is also the only solid ground my team can physically stand on.",
    "continuum_grave_shroud_1_title": "Mass, Sometimes",
    "continuum_grave_shroud_1_desc": "The ship's gravimeter returns errors that look suspiciously like laughter, and then resets itself. We are planting marker flags in soil that only acts like soil until we turn our heads away. A massive structure, perhaps a temple, or a power plant, or somehow both, sits half in and half out of the survey photographs.\n\nWe ordered the drone to photograph it twice. In the second image, the building had sprouted more doors.",
    "continuum_grave_shroud_1_desc_aged": "Desperate pilgrims would have loved this site, right before they were eaten by the object of their love. We are not pilgrims. We are an expedition budget burdened with intense curiosity.\n\nWe have recovered a stone that is only sometimes a stone.",
    "continuum_grave_shroud_2_title": "A Bargain in the Grain",
    "continuum_grave_shroud_2_desc": "Deep microscopy of the soil samples reveals microscopic writing that is not quite writing, penned in a hand that was never biologically a hand. If this spatial event was the result of a dark bargain, the other involved party is still holding their copy of the contract.\n\nWe are not signing anything, not even by leaving a bootprint in the dust.",
    "continuum_grave_shroud_2_desc_aged": "The physical grain of the bedrock repeats a localized phrase we cannot accurately hear. The musicians on our crew swear it sounds almost like a hymn. The captain insists it reads almost like a legal contract.\n\nBoth of them have sensibly put their instruments away.",
    "continuum_grave_shroud_3_title": "Do Not Swim",
    "continuum_grave_shroud_3_desc": "We have packed the maybe-stone, secured the twice-photographed door, and developed a collective headache that will definitely require a ship's chaplain. The formal recommendation is remarkably simple: study this phenomenon from a high orbit.\n\nOnly attempt to colonize this space if you deeply enjoy becoming a cautionary story.",
    "continuum_grave_shroud_3_desc_aged": "Most graves are simply quiet holes dug in the ground. This one is a gaping hole in the fundamental rules of physics. We are leaving a heavy beacon that clearly states, in every language our database holds, that scientific curiosity is not the same thing as offering consent.\n\nThe shifting horizon makes no promises to listen.",
    # --- nanite ---
    "continuum_grave_nanite_cat": "Glass That Used To Be a Biosphere",
    "continuum_grave_nanite_cat_desc": "The surface below us is grey, incredibly fine, and possesses a hungry look even when it is still. We are detecting nanite blooms and L-cluster atmospheric weather. It looks like a metallic tempest that eventually learned how to act like a beach. If there were ever organic people here, they have been reduced to a fine polish on the dust.\n\nIf there is still a dormant queen-signal hidden in those grains, we would highly prefer not to teach it our modern language.",
    "continuum_grave_nanite_cat_desc_aged": "This is a world that was ground down to a single, overriding idea: replicate. The idea itself seems to have grown tired, or perhaps it finally finished its task. The sprawling glass plain is almost pretty from up here.\n\nPretty is how this particular type of grave gets you to take your environmental helmet off.",
    "continuum_grave_nanite_scan_title": "Grey Shore",
    "continuum_grave_nanite_scan_desc": "Our orbital survey of this world has located a vast plain of fine grey dust where a healthy biosphere should naturally exist. The individual grains are microscopic machines, not soil. A scoop taken remotely at the quiet edge of the plain attempted to physically climb the drone's glove before the sample was flash-frozen. There is currently no active control signal broadcasting on the wide band.\n\nThe science officer has logged the open plain and forbidden any further collection. We have one frozen gram in containment, and we will not thaw it near anything we intend to keep alive.",
    "continuum_grave_nanite_scan_desc_aged": "The oldest available charts describe a lush, green world occupying this specific orbit. What actually remains is a massive glass plain composed of dormant machines roughly the size of pollen. Our scientists extracted a single, frozen sample from the planetary tide line and sealed it. The one page of telemetry data we could translate simply stated that the machines were originally built to copy themselves endlessly.\n\nWe have marked the metallic shore and left the interior plain alone. As pretty as the surface looks from orbit, it is definitely not a place to take off a hazmat suit.",
    "continuum_site_nanite": "Quiet Shore",
    "continuum_site_nanite_desc": "We have deployed to the zone where the metallic bloom thins out and something vaguely resembling normal geology begins again. We will work the quiet edge. The deep interior of the plain still registers as an active, dense crowd.",
    "continuum_site_nanite_desc_aged": "We are looking at a literal tide line of dead machines roughly the size of pollen grains. It is visually stunning, provided you remember to keep your helmet on and your gloves sealed.",
    "continuum_grave_nanite_1_title": "Do Not Make a Snowball",
    "continuum_grave_nanite_1_desc": "The first scoop of the dust glitters beautifully in the light. The second scoop immediately tries to scale the outside of the hazmat glove. We flash-froze it, laughed nervously once, and then abruptly stopped laughing.\n\nThis is not standard soil. This is a very patient crowd.",
    "continuum_grave_nanite_1_desc_aged": "Every individual grain is a microscopic library of terrible decisions. We managed to read one data page and immediately closed the book. The page simply said replicate.\n\nWe can assume the next page would have just said it again.",
    "continuum_grave_nanite_2_title": "A War With No Flags",
    "continuum_grave_nanite_2_desc": "There are no torn banners and no records of a heroic last stand. There is only a brutal conversion efficiency that would make a heavy industrial factory weep. If this event was triggered by the Gray Tempest, it was the Tempest operating on a holiday schedule. If it was some local empire's weaponized pet, the pet clearly outlived the masters who fed it.",
    "continuum_grave_nanite_2_desc_aged": "Organic empires always like to sign their weapons of war. Nanites do not. That overwhelming anonymity is their actual signature.\n\nWe have secured a single frozen gram and taken a solemn vow not to thaw it near anything we happen to like.",
    "continuum_grave_nanite_3_title": "Pretty, Logged as Hazard",
    "continuum_grave_nanite_3_desc": "We are leaving a navigational beacon, a firm warning, and a sample crate marked do-not-open in six different scripts. The metallic shore stays pretty in the orbital lenses.\n\nWe are staying in our protective suits.",
    "continuum_grave_nanite_3_desc_aged": "Some archaeological graves politely ask for flowers. This one asks for a heavy lid. We are giving it a lid, a warning flag, and a specific story for the next ship that passes by: the beach down there is not actually a beach.\n\nThe beach is a hungry mouth that simply got tired of chewing.",
    # --- wronghole ---
    "continuum_grave_wronghole_cat": "A Hole With Too Many Planets",
    "continuum_grave_wronghole_cat_desc": "The central star of this system is a black hole, but the local neighborhood does not believe it. The planets still crowd their old, tight orbits, some physically cracked, and some just gravitationally confused. There is no companion supernova remnant, no pretty expansive nebula, and no respectable stellar autobiography left behind.\n\nStars this incredibly tidy do not die like this unless something helps them along.",
    "continuum_grave_wronghole_cat_desc_aged": "The astrometric charts coldly call it a black hole. The surrounding planets call it an active murder scene. Natural stellar collapse does not leave a planetary dinner table set for a warm yellow sun.\n\nSomeone out there ended a star and then left the dishes behind.",
    "continuum_grave_wronghole_scan_title": "Missing Star",
    "continuum_grave_wronghole_scan_desc": "Our deep survey of this system has located a black hole occupying the exact coordinates where the orbits still describe a healthy yellow sun. There is no remnant nebula and no heavy debris shell of the kind a natural gravitational collapse usually leaves behind. The surviving inner worlds still maintain the exact spacing of a family of planets that expected to see the morning light.\n\nThe outer crust of the nearest world is glazed, looking as if the final daylight ended all at once in a flash of heat. We have taken geological samples from that crust. We have decided not to treat the hole as a potential mining site.",
    "continuum_grave_wronghole_scan_desc_aged": "The oldest charts in our database insist a main sequence star once stood here. Our scientists have found a gaping hole instead, and arrayed around it is a set of worlds that still faithfully keep a dead sun's table. Natural stellar collapse does not leave the dishes intact. Someone ended the star and simply left the planets behind in the dark.\n\nWe have recorded the flash glaze on the nearest world and calculated the exact angle of the missing light. The black hole is a galactic crime scene, not a resource claim.",
    "continuum_site_wronghole": "The Witness Orbit",
    "continuum_site_wronghole_desc": "We have established orbit around a surviving world that clearly still remembers having a yellow sky. The central star is gone, but the planetary table is still set. We are going to ask this planet what its final morning actually felt like.",
    "continuum_site_wronghole_desc_aged": "This is a cold planet that patiently kept the table set for a missing host. Natural stellar collapse does not leave the dishes intact. This black hole clearly had some outside help.",
    "continuum_grave_wronghole_1_title": "A Sunburn With No Sun",
    "continuum_grave_wronghole_1_desc": "The outer crust of this witness world is glazed, not by standard weather patterns, but by a final morning that went out like a blown lamp. We found traces of photosynthetic chemistry sulking in the absolute dark. It is waiting for a star that is now just a crushing rule of geometry.\n\nThis was not a slow, creeping death. This was someone flipping a switch.",
    "continuum_grave_wronghole_1_desc_aged": "We have uncovered fragile fossils of leaves that learned how to handle one summer, and then never saw another. The summer ended in a specific physical direction, rather than shifting into a season.\n\nWe are extracting a fossilized leaf. It crumbles under the instruments like a natural law that abruptly got repealed.",
    "continuum_grave_wronghole_2_title": "Helping a Star Die",
    "continuum_grave_wronghole_2_desc": "The local orbital mechanics still describe a standard G-type household. The black hole in the center is a rude guest who murdered the host and arrogantly sat in their favorite chair. A Star-Eater, a Behemoth, or perhaps a dedicated crisis engine: this is the cold physics of helpful murder.\n\nWe are logging the chair. We are not sitting in it.",
    "continuum_grave_wronghole_2_desc_aged": "The spiritual priests would say a god got hungry. The structural engineers would say a massive weapon simply got bored. Both camps would just be dressing up the exact same math. A star was here. Now, a star is not.\n\nThe orbiting planets were certainly not consulted about the change.",
    "continuum_grave_wronghole_3_title": "Leave the Chair Empty",
    "continuum_grave_wronghole_3_desc": "We have gathered enough data: physical samples, an attack vector, and a strict warning that this black hole is not a mining opportunity. It is a mass crime scene. If the entity that did this still lives, it fully knows how to turn off the mornings.\n\nPilot your ships accordingly.",
    "continuum_grave_wronghole_3_desc_aged": "The final instrument our dig uncovered is a child's solar calendar, crudely scratched onto a stone, waiting for a summer solstice that will never be scheduled again. We are packing it between two thick layers of protective foam.\n\nSome evidence is merely geology. Some evidence is pure grief. This black hole happens to be both.",
}
