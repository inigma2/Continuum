"""0.8.7 Ledger: top battle graveyards read from counters Continuum wrote into the save.

The running game records destroyed military hulls on the system. This module
only reads that. It does not write a second ledger file.
"""
import os
import re
import zipfile

MAX_BATTLES = 5
CRISIS_LABEL = {
    "swarm": "the Prethoryn Scourge",
    "extradimensional": "the Unbidden",
    "extradimensional_2": "the Aberrant",
    "extradimensional_3": "the Vehement",
    "ai_empire": "the Contingency",
    "gray_goo": "the Gray Tempest",
    "synth_queen": "the Synth Queen",
    "formless": "the Formless",
}


def _load_gamestate(save_path):
    with zipfile.ZipFile(save_path) as zf:
        return zf.read("gamestate").decode("utf-8", "replace")


def _party_name(cid, names, types):
    """Country names only. A leftover ship id on the flag is not a belligerent."""
    if cid not in names and cid not in types:
        return None, False
    ctype = types.get(cid, "")
    if ctype in CRISIS_LABEL:
        return CRISIS_LABEL[ctype], True
    name = names.get(cid)
    if not name or str(name).isdigit():
        return None, False
    return name, False


def battles_from_galaxy(galaxy_data, countries=None):
    """Systems that recorded destroyed military ships while Continuum was on."""
    names = {}
    types = {}
    for c in countries or []:
        cid = str(c.get("id") or "")
        if not cid:
            continue
        names[cid] = c.get("name") or cid
        types[cid] = str(c.get("type") or "")
    rows = []
    for sys in galaxy_data or []:
        try:
            hulls = int(float(sys.get("continuum_hulls") or 0))
        except (TypeError, ValueError):
            hulls = 0
        if hulls <= 0:
            continue
        victims = []
        killers = []
        crises = []
        vleads = []
        kleads = []
        rulers = []
        goals = []
        for flag in sys.get("flags") or []:
            text = str(flag)
            if text.startswith("continuum_crisis_"):
                label = CRISIS_LABEL.get(text[len("continuum_crisis_") :], text)
                if label not in crises:
                    crises.append(label)
            elif text.startswith("continuum_victim_"):
                cid = text.split("_")[-1]
                label, is_crisis = _party_name(cid, names, types)
                if not label:
                    continue
                if is_crisis and label not in crises:
                    crises.append(label)
                elif label not in victims:
                    victims.append(label)
            elif text.startswith("continuum_killer_"):
                cid = text.split("_")[-1]
                label, is_crisis = _party_name(cid, names, types)
                if not label:
                    continue
                if is_crisis and label not in crises:
                    crises.append(label)
                elif label not in killers:
                    killers.append(label)
            elif text.startswith("continuum_vlead_"):
                vleads.append(text.split("_")[-1])
            elif text.startswith("continuum_klead_"):
                kleads.append(text.split("_")[-1])
            elif text.startswith("continuum_home_ruler_"):
                rulers.append(text.split("_")[-1])
            elif text.startswith("continuum_wargoal_"):
                goals.append(text[len("continuum_wargoal_"):])
        rows.append({
            "system": str(sys.get("id")),
            "size": hulls,
            "hulls": [],
            "victims": victims[:4],
            "killers": killers[:4],
            "crises": crises[:4],
            "admiral_ids": (vleads + kleads)[:6],
            "admirals": [],
            "ruler_ids": rulers[:2],
            "rulers": [],
            "wargoals": goals[:2],
            "empire_best": any(str(f).startswith("continuum_empire_best_") for f in (sys.get("flags") or [])),
            "name": sys.get("name") or "",
            "source": "recorder",
        })
    rows.sort(key=lambda r: (-int(r["size"]), r["system"]))
    return rows


def lookup_leader_names(save_path, ids):
    rows = [{"admiral_ids": [i for i in ids if i], "ruler_ids": []}]
    attach_leader_names(save_path, rows)
    return rows[0].get("_found") or {}


def attach_leader_names(save_path, battles):
    """Fill admiral display names from leader ids stamped on the systems."""
    want = []
    for row in battles or []:
        for lid in list(row.get("admiral_ids") or []) + list(row.get("ruler_ids") or []):
            if lid not in want:
                want.append(lid)
    if not want:
        return battles
    gs = _load_gamestate(save_path)
    found = {}
    for lid in want:
        m = re.search(rf"(?:^|\n)\s*{re.escape(lid)}=\{{", gs)
        if not m:
            continue
        chunk = gs[m.start() : m.start() + 2500]
        if "class=" not in chunk[:900]:
            continue
        name = ""
        custom = re.search(r'custom_name="([^"]+)"', chunk)
        if custom:
            name = custom.group(1)
        if not name:
            first = re.search(r'first_name="([^"]+)"', chunk)
            second = re.search(r'second_name="([^"]+)"', chunk)
            name = " ".join(x.group(1) for x in (first, second) if x).strip()
        if not name:
            plain = re.search(r'\n\s*name="([^"]+)"', chunk)
            if plain and not plain.group(1).startswith(("NAME_", "LEADER_")):
                name = plain.group(1).replace("_", " ")
        if name:
            found[lid] = name
    for row in battles or []:
        row["admirals"] = [found[i] for i in (row.get("admiral_ids") or []) if i in found]
        row["rulers"] = [found[i] for i in (row.get("ruler_ids") or []) if i in found]
        row["_found"] = found
    return battles


def top_battles(rows):
    """Galaxy top five, plus each empire's own largest fight if it is not already in that five."""
    ranked = sorted(rows or [], key=lambda r: (-int(r.get("size") or 0), str(r.get("system"))))
    chosen = []
    seen = set()
    for row in ranked[:MAX_BATTLES]:
        chosen.append(row)
        seen.add(str(row.get("system")))
    for row in ranked:
        sid = str(row.get("system"))
        if row.get("empire_best") and sid not in seen:
            chosen.append(row)
            seen.add(sid)
    for i, row in enumerate(chosen):
        row["idx"] = i
    return chosen


def apply_battle_flags(battles, extra_flags):
    flags = extra_flags if extra_flags is not None else {}
    for row in battles or []:
        sid = str(row["system"])
        flag = f"continuum_battle_{row['idx']}"
        lst = flags.setdefault(sid, [])
        if flag not in lst:
            lst.append(flag)
    return flags


def _battle_desc(row, aged=False):
    row = row or {}
    sides = []
    for name in list(row.get("crises") or []) + list(row.get("victims") or []) + list(row.get("killers") or []):
        if name and name not in sides:
            sides.append(str(name))
    admirals = [str(a) for a in (row.get("admirals") or []) if a]
    if not sides and not admirals:
        if aged:
            return (
                "Old charts call this stretch of dark a graveyard. Our scientists cannot name the two navies, "
                "but the wrecks still lie in two lines, and the metal scans as military hull armor rather than raw asteroid ore. "
                "Whoever fought here abandoned the dead ships where they fell.\\n\\n"
                "We have extracted a salvage from the field. The rest of the graveyard remains drifting "
                "for the next ship that bothers to look."
            )
        return (
            "Our science ship has finished a survey of the wreck field. The hulks still answer with two "
            "identification codes, and neither navy is here to claim them. The metal is cold, but the line "
            "of the battle is still readable from the way the wrecks face one another.\\n\\n"
            "We have stripped what the field will safely give. The alloys and minerals stay in this system "
            "for any station that comes after us."
        )
    named = sides[:4]
    if len(named) == 0:
        who = "two navies whose names are gone"
    elif len(named) == 1:
        who = named[0]
    elif len(named) == 2:
        who = f"{named[0]} and {named[1]}"
    else:
        who = ", ".join(named[:-1]) + ", and " + named[-1]
    command = ""
    rulers = [str(a) for a in (row.get("rulers") or []) if a]
    if rulers:
        command += " A recovered dispatch is addressed to " + ", ".join(rulers[:2]) + "."
    if admirals:
        command += " It is signed by " + ", ".join(admirals[:2]) + "."
    goals = [str(g).replace("_", " ") for g in (row.get("wargoals") or []) if g]
    if aged:
        goal_bit = ""
        if goals:
            goal_bit = ", and they record the goal of the fight as " + goals[0]
        return (
            f"Our oldest charts still name the fleets that died in this dark as {who}{goal_bit}.{command} "
            "Our scientists can still tell the wrecks apart by the alloy of their hulls, and the lines they hold "
            "are the lines of a battle rather than a scrap yard.\\n\\n"
            "We have taken a salvage from the field and left the rest of the graveyard drifting where it fell. "
            "The next survey ship to come this way will find those same lines of metal still waiting in the dark."
        )
    if goals:
        command += " The war was logged as " + goals[0] + "."
    return (
        f"Our science ship has completed a survey of the wreck field. The hulks still carry the marks of {who}.{command} "
        "Neither navy is here to claim the dead. The metal is cold, but the line of the battle is still readable "
        "from the way the wrecks face one another.\\n\\n"
        "We have stripped what salvage the field will safely yield. The alloys and minerals will stay in this system "
        "for any station that comes after us."
    )


def _promote_best(who):
    """Point this empire at the current system when it is their largest fight so far."""
    return f"""if = {{
					limit = {{
						exists = {who}
						OR = {{
							{who} = {{ NOT = {{ check_variable = {{ which = continuum_best_hulls value > 0 }} }} }}
							check_variable = {{ which = continuum_hulls value > {who}.continuum_best_hulls }}
						}}
					}}
					every_system = {{
						limit = {{ has_star_flag = continuum_empire_best_@{who} }}
						remove_star_flag = continuum_empire_best_@{who}
					}}
					every_system = {{
						limit = {{ {who} = {{ has_country_flag = continuum_best_at_@prev }} }}
						{who} = {{ remove_country_flag = continuum_best_at_@prev }}
					}}
					set_star_flag = continuum_empire_best_@{who}
					{who} = {{
						set_variable = {{ which = continuum_best_hulls value = prev.continuum_hulls }}
						set_country_flag = continuum_best_at_@prev
					}}
				}}"""


def _crisis_flag_block(scope):
    lines = []
    for key in CRISIS_LABEL:
        lines.append(
            f"""if = {{
					limit = {{ {scope} = {{ is_country_type = {key} }} }}
					set_star_flag = continuum_crisis_{key}
				}}"""
        )
    return "\n\t\t\t\t".join(lines)


def write_watch_files(mod_root):
    """Recorder must exist during the Pre game, before any parse."""
    events = os.path.join(mod_root, "events")
    actions = os.path.join(mod_root, "common", "on_actions")
    os.makedirs(events, exist_ok=True)
    os.makedirs(actions, exist_ok=True)
    with open(os.path.join(actions, "zz_continuum_ledger.txt"), "w", encoding="utf-8") as fh:
        fh.write("""# Counts military hulls lost in each system. Survives after debris expires.
on_ship_destroyed_victim = {
	events = {
		continuum_ledger.10
	}
}
on_yearly_pulse = {
	events = {
		continuum_ledger.11
	}
}
on_country_destroyed = {
	events = {
		continuum_ledger.20
	}
}
on_destroy_planet_with_PLANET_KILLER_CRACKER = {
	events = {
		continuum_ledger.30
	}
}
on_destroy_planet_with_PLANET_KILLER_SHIELDER = {
	events = {
		continuum_ledger.31
	}
}
on_destroy_planet_with_PLANET_KILLER_NEUTRON = {
	events = {
		continuum_ledger.32
	}
}
on_destroy_planet_with_PLANET_KILLER_GODRAY = {
	events = {
		continuum_ledger.33
	}
}
""")
    with open(os.path.join(events, "continuum_ledger_watch.txt"), "w", encoding="utf-8") as fh:
        fh.write("""namespace = continuum_ledger

# This = owner of the destroyed ship. From = killer. FromFrom = destroyed ship.
country_event = {
	id = continuum_ledger.10
	hide_window = yes
	is_triggered_only = yes
	trigger = {
		exists = fromfrom
		fromfrom = {
			is_space_fauna = no
			OR = {
				is_ship_class = shipclass_military
				is_ship_class = shipclass_military_special
				is_ship_class = shipclass_starbase
			}
		}
	}
	immediate = {
		fromfrom = {
			fleet = {
				if = {
					limit = { exists = leader }
					leader = { save_event_target_as = cont_v_leader }
				}
			}
			solar_system = {
				change_variable = { which = continuum_hulls value = 1 }
				set_variable = { which = continuum_rank_worse value = 0 }
				every_system = {
					limit = {
						NOT = { is_same_value = prev }
						check_variable = { which = continuum_hulls value > 0 }
					}
					set_variable = { which = continuum_their_hulls value = prev.continuum_hulls }
					if = {
						limit = { check_variable = { which = continuum_hulls value > continuum_their_hulls } }
						prev = { change_variable = { which = continuum_rank_worse value = 1 } }
					}
				}
				""" + _promote_best("root") + """
				if = {
					limit = {
						OR = {
							check_variable = { which = continuum_rank_worse value < 5 }
							has_star_flag = continuum_empire_best_@root
						}
					}
					set_star_flag = continuum_victim_@root
					if = {
						limit = { exists = event_target:cont_v_leader }
						set_star_flag = continuum_vlead_@event_target:cont_v_leader
					}
					if = {
						limit = {
							check_variable = { which = continuum_hulls value >= 8 }
							NOT = { has_star_flag = continuum_ruler_set }
						}
						set_star_flag = continuum_ruler_set
						root = {
							if = {
								limit = { exists = ruler }
								ruler = { save_event_target_as = cont_home_ruler }
							}
						}
						if = {
							limit = { exists = event_target:cont_home_ruler }
							set_star_flag = continuum_home_ruler_@event_target:cont_home_ruler
						}
						if = {
							limit = { root = { any_war = { using_war_goal = { type = wg_conquest owner = root } } } }
							set_star_flag = continuum_wargoal_conquest
						}
						if = {
							limit = { root = { any_war = { using_war_goal = { type = wg_humiliation owner = root } } } }
							set_star_flag = continuum_wargoal_humiliation
						}
						if = {
							limit = { root = { any_war = { using_war_goal = { type = wg_subjugation owner = root } } } }
							set_star_flag = continuum_wargoal_subjugation
						}
						if = {
							limit = { root = { any_war = { using_war_goal = { type = wg_end_threat owner = root } } } }
							set_star_flag = continuum_wargoal_end_threat
						}
					}
					""" + _crisis_flag_block("root") + """
				}
			}
		}
		if = {
			limit = { exists = fromfromfrom }
			fromfromfrom = {
				fleet = {
					owner = { save_event_target_as = cont_killer_country }
					if = {
						limit = { exists = leader }
						leader = { save_event_target_as = cont_k_leader }
					}
				}
				solar_system = {
					""" + _promote_best("event_target:cont_killer_country") + """
					if = {
						limit = {
							OR = {
								check_variable = { which = continuum_rank_worse value < 5 }
								has_star_flag = continuum_empire_best_@root
								has_star_flag = continuum_empire_best_@event_target:cont_killer_country
							}
						}
						if = {
							limit = { exists = event_target:cont_killer_country }
							set_star_flag = continuum_killer_@event_target:cont_killer_country
						}
						if = {
							limit = { exists = event_target:cont_k_leader }
							set_star_flag = continuum_klead_@event_target:cont_k_leader
						}
						""" + _crisis_flag_block("event_target:cont_killer_country") + """
					}
				}
			}
		}
	}
}

# Once a year, drop story flags from every battlefield that is no longer in the top five.
# The hull count stays, so a later fight can climb back into the five.
event = {
	id = continuum_ledger.11
	hide_window = yes
	is_triggered_only = yes
	immediate = {
		every_system = {
			limit = { check_variable = { which = continuum_hulls value > 0 } }
			set_variable = { which = continuum_rank_worse value = 0 }
			every_system = {
				limit = {
					NOT = { is_same_value = prev }
					check_variable = { which = continuum_hulls value > 0 }
				}
				set_variable = { which = continuum_their_hulls value = prev.continuum_hulls }
				if = {
					limit = { check_variable = { which = continuum_hulls value > continuum_their_hulls } }
					prev = { change_variable = { which = continuum_rank_worse value = 1 } }
				}
			}
		}
		every_system = {
			limit = { check_variable = { which = continuum_hulls value > 0 } }
			set_variable = { which = continuum_still_best value = 0 }
			every_country = {
				limit = { has_country_flag = continuum_best_at_@prev }
				prev = { change_variable = { which = continuum_still_best value = 1 } }
			}
			if = {
				limit = {
					check_variable = { which = continuum_rank_worse value >= 5 }
					check_variable = { which = continuum_still_best value = 0 }
				}
			remove_star_flag = continuum_ruler_set
			remove_star_flag = continuum_wargoal_conquest
			remove_star_flag = continuum_wargoal_humiliation
			remove_star_flag = continuum_wargoal_subjugation
			remove_star_flag = continuum_wargoal_end_threat
			remove_star_flag = continuum_crisis_swarm
			remove_star_flag = continuum_crisis_extradimensional
			remove_star_flag = continuum_crisis_extradimensional_2
			remove_star_flag = continuum_crisis_extradimensional_3
			remove_star_flag = continuum_crisis_ai_empire
			remove_star_flag = continuum_crisis_gray_goo
			remove_star_flag = continuum_crisis_synth_queen
			remove_star_flag = continuum_crisis_formless
			every_country = {
				prev = {
					remove_star_flag = continuum_victim_@prev
					remove_star_flag = continuum_killer_@prev
				}
				every_owned_leader = {
					save_event_target_as = cont_clear_leader
					prevprev = {
						remove_star_flag = continuum_vlead_@event_target:cont_clear_leader
						remove_star_flag = continuum_klead_@event_target:cont_clear_leader
						remove_star_flag = continuum_home_ruler_@event_target:cont_clear_leader
					}
				}
			}
			}
		}
	}
}

# This = destroyed country. Stamp its ruler on the capital for a later dig.
country_event = {
	id = continuum_ledger.20
	hide_window = yes
	is_triggered_only = yes
	trigger = {
		OR = {
			is_country_type = default
			is_country_type = fallen_empire
			is_country_type = awakened_fallen_empire
			is_country_type = swarm
			is_country_type = extradimensional
			is_country_type = extradimensional_2
			is_country_type = extradimensional_3
			is_country_type = ai_empire
			is_country_type = gray_goo
			is_country_type = synth_queen
			is_country_type = formless
		}
	}
	immediate = {
		if = { limit = { is_country_type = swarm } set_global_flag = continuum_crisis_seen_swarm }
		if = { limit = { is_country_type = extradimensional } set_global_flag = continuum_crisis_seen_extradimensional }
		if = { limit = { is_country_type = extradimensional_2 } set_global_flag = continuum_crisis_seen_extradimensional_2 }
		if = { limit = { is_country_type = extradimensional_3 } set_global_flag = continuum_crisis_seen_extradimensional_3 }
		if = { limit = { is_country_type = ai_empire } set_global_flag = continuum_crisis_seen_ai_empire }
		if = { limit = { is_country_type = gray_goo } set_global_flag = continuum_crisis_seen_gray_goo }
		if = { limit = { is_country_type = synth_queen } set_global_flag = continuum_crisis_seen_synth_queen }
		if = { limit = { is_country_type = formless } set_global_flag = continuum_crisis_seen_formless }
		if = {
			limit = {
				exists = ruler
				exists = capital_scope
				OR = {
					is_country_type = default
					is_country_type = fallen_empire
					is_country_type = awakened_fallen_empire
				}
			}
			ruler = { save_event_target_as = cont_dead_ruler }
			capital_scope = {
				set_planet_flag = continuum_fallen_ruler_@event_target:cont_dead_ruler
			}
		}
	}
}

# This = planet fired upon. From = fleet that fired.
planet_event = {
	id = continuum_ledger.30
	hide_window = yes
	is_triggered_only = yes
	immediate = {
		set_planet_flag = continuum_pk_cracker
		from = {
			owner = {
				save_event_target_as = cont_pk_owner
				if = {
					limit = { exists = ruler }
					ruler = { save_event_target_as = cont_pk_ruler }
				}
			}
		}
		if = {
			limit = { exists = event_target:cont_pk_ruler }
			set_planet_flag = continuum_pk_ruler_@event_target:cont_pk_ruler
		}
	}
}
planet_event = {
	id = continuum_ledger.31
	hide_window = yes
	is_triggered_only = yes
	immediate = {
		set_planet_flag = continuum_pk_shielder
		from = {
			owner = {
				if = {
					limit = { exists = ruler }
					ruler = { save_event_target_as = cont_pk_ruler }
				}
			}
		}
		if = {
			limit = { exists = event_target:cont_pk_ruler }
			set_planet_flag = continuum_pk_ruler_@event_target:cont_pk_ruler
		}
	}
}
planet_event = {
	id = continuum_ledger.32
	hide_window = yes
	is_triggered_only = yes
	immediate = {
		set_planet_flag = continuum_pk_neutron
		from = {
			owner = {
				if = {
					limit = { exists = ruler }
					ruler = { save_event_target_as = cont_pk_ruler }
				}
			}
		}
		if = {
			limit = { exists = event_target:cont_pk_ruler }
			set_planet_flag = continuum_pk_ruler_@event_target:cont_pk_ruler
		}
	}
}
planet_event = {
	id = continuum_ledger.33
	hide_window = yes
	is_triggered_only = yes
	immediate = {
		set_planet_flag = continuum_pk_godray
		from = {
			owner = {
				if = {
					limit = { exists = ruler }
					ruler = { save_event_target_as = cont_pk_ruler }
				}
			}
		}
		if = {
			limit = { exists = event_target:cont_pk_ruler }
			set_planet_flag = continuum_pk_ruler_@event_target:cont_pk_ruler
		}
	}
}
""")


def _battle_desc_blocks(battles):
    lines = []
    for i, row in enumerate(battles or []):
        lines.append(
            f"""desc = {{
		trigger = {{
			has_global_flag = continuum_aged
			solar_system = {{ has_star_flag = continuum_battle_{i} }}
		}}
		text = continuum_battle_{i}_desc_aged
	}}
	desc = {{
		trigger = {{
			solar_system = {{ has_star_flag = continuum_battle_{i} }}
		}}
		text = continuum_battle_{i}_desc
	}}"""
        )
    return "\n\t".join(lines)


def write_ledger_files(mod_root, loc_entries, battles=None):
    write_watch_files(mod_root)
    common = os.path.join(mod_root, "common", "anomalies")
    events = os.path.join(mod_root, "events")
    os.makedirs(common, exist_ok=True)
    os.makedirs(events, exist_ok=True)
    with open(os.path.join(common, "continuum_ledger.txt"), "w", encoding="utf-8") as fh:
        fh.write("""# Continuum 0.8.7 battle graveyards. Seeded by event, not by survey chance.
continuum_battle_cat = {
	desc = continuum_battle_cat_desc
	picture = GFX_evt_ship_in_orbit
	level = 4
	spawn_chance = {
		base = 0
	}
	on_success = continuum_ledger.1
}
""")
    bits = ["""namespace = continuum_ledger

event = {
	id = continuum_ledger.0
	hide_window = yes
	is_triggered_only = yes
	immediate = {
"""]
    for i in range(len(battles or [])):
        bits.append(f"""		every_system = {{
			limit = {{ has_star_flag = continuum_battle_{i} }}
			random_system_planet = {{
				limit = {{ is_star = yes }}
				add_anomaly = {{ category = continuum_battle_cat }}
			}}
		}}
""")
    bits.append("""	}
}

ship_event = {
	id = continuum_ledger.1
	title = continuum_battle_title
	""" + _battle_desc_blocks(battles) + """
	desc = continuum_battle_desc
	picture = GFX_evt_ship_in_orbit
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {
		name = continuum_battle_ok
		FROM = {
			add_deposit = d_alloys_5
			add_deposit = d_minerals_5
		}
		hidden_effect = {
			from = {
				create_ambient_object = {
					type = derelict_ship_object
					location = this
				}
				last_created_ambient_object = {
					save_event_target_as = continuum_wreck
				}
			}
			event_target:continuum_wreck = {
				enable_special_project = {
					name = CONTINUUM_BATTLE_SALVAGE
					owner = root
					location = this
				}
			}
		}
	}
}

ship_event = {
	id = continuum_ledger.2
	title = continuum_battle_salvage_title
	desc = continuum_battle_salvage_desc
	picture = GFX_evt_ship_in_orbit
	show_sound = event_ship_bridge
	location = FROM
	is_triggered_only = yes
	option = {
		name = continuum_battle_salvage_ok
		owner = {
			create_fleet = {
				settings = {
					spawn_debris = no
					can_upgrade = no
					can_change_composition = no
					uses_naval_capacity = no
				}
				effect = {
					set_owner = prev
					create_ship = {
						name = "Salvaged Hull"
						random_existing_design = corvette
						prefix = no
					}
					set_location = root
				}
			}
		}
	}
}
""")
    with open(os.path.join(events, "continuum_ledger_events.txt"), "w", encoding="utf-8") as fh:
        fh.write("".join(bits))
    proj = os.path.join(mod_root, "common", "special_projects")
    os.makedirs(proj, exist_ok=True)
    with open(os.path.join(proj, "continuum_battle_salvage.txt"), "w", encoding="utf-8") as fh:
        fh.write("""special_project = {
	key = "CONTINUUM_BATTLE_SALVAGE"
	days_to_research = 30
	tech_department = engineering_technology
	timelimit = -1
	picture = GFX_evt_ship_in_orbit
	icon = "gfx/interface/icons/situation_log/situation_log_debris.dds"
	event_scope = ship_event
	requirements = {
		shipclass_science_ship = 1
		leader = scientist
	}
	on_success = {
		ship_event = { id = continuum_ledger.2 }
	}
}
""")
    loc_entries.update({
        "continuum_battle_cat": "Debris Field",
        "continuum_battle_cat_desc": "A field of dead warships still holds a rough line. Two navies died here, and nobody came back for the metal.",
        "continuum_battle_title": "Two Transponders",
        "continuum_battle_desc": _battle_desc((battles or [None])[0], aged=False),
        "continuum_battle_desc_aged": _battle_desc((battles or [None])[0], aged=True),
        "continuum_battle_ok": "Strip the field.",
        "continuum_battle_salvage_title": "The Hull Still Answers",
        "continuum_battle_salvage_desc": "The wreck's transponder accepted a new owner. One corvette is intact enough to bring home. It will not take a refit, and it does not count against the naval limit.",
        "continuum_battle_salvage_ok": "Bring the hull home.",
        "CONTINUUM_BATTLE_SALVAGE": "Salvage the Derelict",
        "CONTINUUM_BATTLE_SALVAGE_DESC": "A dead hull from the battle still holds together. A science ship can bring it in.",
    })
    for i, row in enumerate(battles or []):
        loc_entries[f"continuum_battle_{i}_desc"] = _battle_desc(row, aged=False)
        loc_entries[f"continuum_battle_{i}_desc_aged"] = _battle_desc(row, aged=True)
