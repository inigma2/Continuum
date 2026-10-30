"""0.8.9 Sleepers: one or two Pre Fallens become Aged precursor digs."""
import os


def _walk(sys):
    out = []
    if not sys:
        return out
    root = sys.get("hierarchy_root") or sys
    q = list(root.get("children") or [])
    while q:
        b = q.pop(0)
        out.append(b)
        q.extend(b.get("children") or [])
    return out


_HAB = (
    "pc_continental", "pc_tropical", "pc_ocean", "pc_arid", "pc_desert",
    "pc_tundra", "pc_arctic", "pc_alpine", "pc_savannah", "pc_gaia",
    "pc_relic", "pc_nuked", "pc_hive", "pc_machine",
)


def _usable(body):
    pc = str(body.get("planet_class") or "")
    if not pc or "star" in pc or "asteroid" in pc:
        return False
    flags = " ".join(str(f) for f in (body.get("flags") or []))
    if "continuum_gsite" in flags or "continuum_grave" in flags:
        return False
    return True


def _best_body(bodies, skip_id=""):
    ranked = []
    for b in bodies:
        if not _usable(b) or str(b.get("id")) == str(skip_id):
            continue
        pc = str(b.get("planet_class") or "")
        ranked.append((0 if pc in _HAB else 1, str(b.get("id")), b))
    ranked.sort()
    return ranked[0][2] if ranked else None


def bind(plan, galaxy):
    sys_by = {str(s.get("id")): s for s in galaxy or []}
    digs = {}
    clues = {}
    for cid in plan.get("sleeper_ids") or []:
        home = str((plan.get("sleeper_home") or {}).get(cid) or "")
        picked = _best_body(_walk(sys_by.get(home)))
        dig = str(picked.get("id")) if picked else ""
        digs[cid] = dig
        found = []
        for sid in (plan.get("sleeper_cores") or {}).get(cid) or []:
            if str(sid) == home:
                continue
            clue = _best_body(_walk(sys_by.get(str(sid))), dig)
            if clue:
                found.append(str(clue.get("id")))
            if len(found) >= 2:
                break
        clues[cid] = found
    plan["sleeper_dig"] = digs
    plan["sleeper_clues"] = clues


def stamp(plan, star_flags, planet_flags):
    for i, cid in enumerate(plan.get("sleeper_ids") or []):
        for sid in (plan.get("sleeper_cores") or {}).get(cid) or []:
            star_flags.setdefault(str(sid), []).append(f"continuum_sleeper_{i}")
        home = str((plan.get("sleeper_home") or {}).get(cid) or "")
        if home:
            star_flags.setdefault(home, []).append(f"continuum_sleeper_{i}_capital")
        dig = (plan.get("sleeper_dig") or {}).get(cid) or ""
        if dig:
            planet_flags.setdefault(dig, []).append(f"continuum_sleeper_{i}_home")
        for n, pid in enumerate((plan.get("sleeper_clues") or {}).get(cid) or []):
            planet_flags.setdefault(pid, []).append(f"continuum_sleeper_{i}_clue_{n}")


def _voice(emp):
    ethics = " ".join(emp.get("ethics") or [])
    if "xenophobe" in ethics or "militarist" in ethics:
        return "march"
    return "catalog"


def write_files(mod_root, loc, plan):
    events = os.path.join(mod_root, "events")
    anom = os.path.join(mod_root, "common", "anomalies")
    sites = os.path.join(mod_root, "common", "archaeological_site_types")
    relics = os.path.join(mod_root, "common", "relics")
    for d in (events, anom, sites, relics):
        os.makedirs(d, exist_ok=True)
    emps = plan.get("sleeper_emps") or {}
    ids = list(plan.get("sleeper_ids") or [])
    cat_bits = []
    site_bits = []
    ev = ["namespace = continuum_sleeper\n", """event = {
	id = continuum_sleeper.0
	hide_window = yes
	is_triggered_only = yes
	trigger = { has_global_flag = continuum_aged }
	immediate = {
"""]
    # ship_event / fleet_event must be siblings of event .0, not nested in immediate.
    outside = []
    for i, cid in enumerate(ids):
        emp = emps.get(cid) or {}
        name = emp.get("name") or "a fallen empire"
        voice = _voice(emp)
        loc[f"continuum_sleeper_{i}_name"] = name
        if voice == "march":
            loc[f"continuum_sleeper_{i}_site"] = "The Empty March"
            loc[f"continuum_sleeper_{i}_site_desc"] = (
                f"From high orbit, the capital of {name} reads as a fortress that outlived its garrison. "
                "Every gun mount our scientists can resolve was laid on the lanes, not aimed inward at the palace, "
                "and the barracks still outnumber the archives.\\n\\n"
                "The guest quarters that can still be picked out from orbit are broken open, and nothing on the scan "
                "suggests this empire ever welcomed a neighbor inside its walls. "
                "The excavation will begin in the fused armor of the capital itself."
            )
            chapters = [
                ("Guns Trained on the Lane",
                 "Our scientists have cut into the capital along a trench of fused armor. The walls faced outward. "
                 "On the inner face, the team logged plated warnings, repeated in several dead scripts, that ordered foreign hulls to turn back or burn.\\n\\n"
                 f"The empire of {name} built this world as a gate and lived behind it. What remains are holding cells, firing ranges, and a parade ground vitrified by its own last volley."),
                ("Ranges Aimed at Empty Space",
                 "The second trench opened a firing range whose targets were painted star charts. The charts showed neighboring systems, each marked as forbidden. "
                 "No round in the racks had been fired at those paintings. The guns had been saved for real hulls.\\n\\n"
                 "Duty slates name captains who held the range after the outer forts went silent. The last slate is a bearing and a hold order. They died waiting for an enemy that had already won."),
                ("The Armory That Stayed Shut",
                 "In the deepest vault, the excavation found an armory door that was still sealed. Our scientists tested the seam. It did not yield. "
                 "A plaque in the lintel came free. It named this capital and charged the garrison to keep the armory shut.\\n\\n"
                 f"The empire of {name} ended as a closed fortress. The plaque is ours. The armory stays shut."),
            ]
            clues = [
                ("Warning Buoy on a Closed Lane",
                 f"Our scientists boarded a buoy that still repeats a warning in a dead script. The message names {name} and orders traffic to turn back. "
                 "The buoy's return channel still carried a spectral type and a bearing for the capital."),
                ("Hulks at a Forbidden Star",
                 f"A field of hulks shared one transponder prefix. The ships had been turned away from a single star and had died on that heading. "
                 f"The prefix matches the garrison rolls of {name}. The heading is enough to name the capital system."),
            ]
        else:
            loc[f"continuum_sleeper_{i}_site"] = "The Closed Catalog"
            loc[f"continuum_sleeper_{i}_site_desc"] = (
                f"From high orbit, the capital of {name} is a muted disk under a broken ring of stations. "
                "The ring was a gallery of docks and reading halls, and our scientists can still trace the tram lines between the dark museum shells. "
                "The cities on the night side keep their street grids, but no light answers.\\n\\n"
                "The empire treated the galaxy as a collection, and this world as the drawer that held it. "
                "The excavation will begin in the capital itself."
            )
            chapters = [
                ("The Fallen Reading Hall",
                 f"Our scientists opened the first trench in a hall that had fallen from the orbital ring onto the capital of {name}. "
                 "Shelving ran for kilometers in a metal that had grown tired rather than rusted. Labels survived in a script the excavation can scan but cannot yet speak.\\n\\n"
                 "The shelves had held star charts, grammars, and sealed copies of laws. The arrangement was a filing system, not a shrine. "
                 "The empire had been vast enough to gather this, and careful enough to catalog it. It had not been careful enough to last."),
                ("Rosters That Never Signed Out",
                 "A service tunnel led from the fallen hall into the lower stacks. Duty rosters still hung on the walls. "
                 "The last names had signed in and had never signed out. Ration marks showed a staff that kept working after the docks above them went quiet.\\n\\n"
                 "The excavation recovered a maintenance log, not a battle plan. The entries grew shorter. The final line was a shelf number, not a farewell."),
                ("The Seal on the Last Drawer",
                 f"In the deepest vault, the excavation found the archive the upper halls had only indexed. The door was one slab, marked with the catalog number of {name}, "
                 "as if the rulers had filed themselves. Our scientists tested the seam. It did not yield.\\n\\n"
                 "A plaque in the lintel came free. It carried a star-fix for this capital and a charge to keep the record and not to spend it. "
                 "The plaque is a relic of that charge. The archive stayed sealed."),
            ]
            clues = [
                ("A Chart Left in a Dead Relay",
                 f"Our scientists boarded a dead relay off the lane. Inside, a chart cabinet held a loan record. The station had borrowed star maps from {name} and never sent them home. "
                 "The maps are ash. The return address is not. A bearing and a catalog number for the capital ring survived in the frame."),
                ("A Copy That Named Its Source",
                 "A ruined colony had streets numbered like volumes in a library. In the civic pit our scientists found a cracked copy of a star catalog. "
                 f"A credit line under the break named {name} and the ring where the original was kept. The entries were gone. The credit was enough."),
            ]
        for n, (title, body) in enumerate(chapters, 1):
            loc[f"continuum_sleeper_{i}_ch{n}"] = title
            loc[f"continuum_sleeper_{i}_ch{n}_desc"] = body
        loc[f"continuum_sleeper_{i}_ok"] = "Continue the excavation"
        for n, (title, body) in enumerate(clues):
            loc[f"continuum_sleeper_{i}_clue_{n}"] = title
            loc[f"continuum_sleeper_clue_{i}_{n}"] = title
            loc[f"continuum_sleeper_{i}_clue_{n}_desc"] = body + "\\n\\nThe find names the capital system. Our scientists logged the location."
            loc[f"continuum_sleeper_{i}_clue_{n}_ok"] = "Log the bearing."
            cat_bits.append(f"""continuum_sleeper_clue_{i}_{n} = {{
	desc = continuum_sleeper_{i}_clue_{n}_desc
	picture = GFX_evt_archaeological_dig
	level = 2
	spawn_chance = {{ base = 0 }}
	on_success = continuum_sleeper.{20 + i * 2 + n}
}}
""")
            ev.append(f"""		every_galaxy_planet = {{
			limit = {{ has_planet_flag = continuum_sleeper_{i}_clue_{n} }}
			add_anomaly = {{ category = continuum_sleeper_clue_{i}_{n} }}
		}}
""")
            outside.append(f"""ship_event = {{
	id = continuum_sleeper.{20 + i * 2 + n}
	title = continuum_sleeper_{i}_clue_{n}
	desc = continuum_sleeper_{i}_clue_{n}_desc
	picture = GFX_evt_archaeological_dig
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {{ name = continuum_sleeper_{i}_clue_{n}_ok }}
}}
""")
        if (plan.get("sleeper_dig") or {}).get(cid):
            site_bits.append(f"""continuum_sleeper_site_{i} = {{
	desc = continuum_sleeper_{i}_site_desc
	picture = GFX_evt_archaeological_dig
	stages = 3
	weight = 0
	allow = {{ exists = leader leader = {{ leader_class = scientist }} }}
	visible = {{ default_site_visible_trigger = yes }}
	stage = {{ difficulty = 2 icon = GFX_archaeology_icon_loot event = continuum_sleeper.{10 + i * 3 + 1} }}
	stage = {{ difficulty = 3 icon = GFX_archaeology_icon_loot event = continuum_sleeper.{10 + i * 3 + 2} }}
	stage = {{ difficulty = 4 icon = GFX_archaeology_icon_success event = continuum_sleeper.{10 + i * 3 + 3} }}
	on_roll_failed = {{
		from = {{
			standard_archaeological_site_on_roll_failed = {{ RANDOM_EVENTS = no_events }}
		}}
	}}
}}
""")
            ev.append(f"""		every_galaxy_planet = {{
			limit = {{ has_planet_flag = continuum_sleeper_{i}_home }}
			create_archaeological_site = continuum_sleeper_site_{i}
		}}
""")
            for n, (title, _body) in enumerate(chapters, 1):
                reward = ""
                if n == 3:
                    deposit = "d_engineering_5" if voice == "march" else "d_society_5"
                    relic = ""
                    if i == 0:
                        # 4.5 grants a relic with add_relic. give_relic is not an effect.
                        relic = """
			if = {
				limit = { has_ancrel = yes NOT = { has_relic = r_continuum_sleeper } }
				add_relic = r_continuum_sleeper
			}"""
                    reward = f"""
		owner = {{
			add_monthly_resource_mult = {{ resource = society_research value = 10 min = 200 max = 400 }}
			add_resource = {{ alloys = 80 }}{relic}
		}}
		from = {{ add_deposit = {deposit} }}"""
                elif n == 1:
                    reward = """
		owner = {
			add_monthly_resource_mult = { resource = society_research value = 4 min = 80 max = 150 }
		}"""
                else:
                    reward = """
		owner = { add_resource = { minerals = 150 } }"""
                outside.append(f"""fleet_event = {{
	id = continuum_sleeper.{10 + i * 3 + n}
	title = continuum_sleeper_{i}_ch{n}
	desc = continuum_sleeper_{i}_ch{n}_desc
	picture = GFX_evt_archaeological_dig
	show_sound = event_ship_bridge
	archaeology = yes
	location = from
	is_triggered_only = yes
	option = {{
		name = continuum_sleeper_{i}_ok{reward}
	}}
}}
""")
    ev.append("	}\n}\n")
    ev.extend(outside)
    with open(os.path.join(anom, "continuum_sleeper_anomalies.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(cat_bits))
    with open(os.path.join(sites, "continuum_sleeper_sites.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(site_bits))
    with open(os.path.join(events, "continuum_sleeper_events.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(ev))
    with open(os.path.join(relics, "continuum_sleeper_relics.txt"), "w", encoding="utf-8") as fh:
        fh.write("""r_continuum_sleeper = {
	activation_duration = 3600
	portrait = "GFX_relic_galatron"
	sound = "relic_activation_galatron"
	score = 500
	resources = {
		category = relics
		cost = { energy = 1000 }
	}
	ai_weight = { weight = 0 }
	triggered_country_modifier = {
		potential = { always = yes }
		country_society_research_produces_mult = 0.05
	}
	active_effect = {
		add_monthly_resource_mult = { resource = society_research value = 6 min = 100 max = 300 }
	}
	possible = { is_ai = no }
}
""")
    loc["r_continuum_sleeper"] = "Sealed Catalog"
    loc["r_continuum_sleeper_desc"] = "A plaque from a fallen capital. It charges the holder to keep the record and not to spend it."
