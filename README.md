# Continuum

A Stellaris galaxy converter — one galaxy to rule them all.

Turn a finished (or mid-game) save into a playable **static galaxy** for a new game. The Python parser reads your `.sav` and writes the imported galaxy into the Continuum mod folder.

- **Current:** 0.7.5 for Stellaris **4.4.*** (Pegasus) — **Continuum Present**
- Workshop: https://steamcommunity.com/sharedfiles/filedetails/?id=3554276594
- Roadmap: https://steamcommunity.com/workshop/filedetails/discussion/3554276594/596284386694022138/

## Continuum Present (0.7.5)

New Game on **Continuum Present** is a **2200** start in the imported galaxy. You are always a **new polity**. The old save’s empires come back as living NPCs.

Restores, as far as Stellaris scripts allow:

- Systems, hyperlanes, stars, planets, moons, asteroids, belts, nebulas
- Megastructures, natural wormholes, shroud tunnels, L-Cluster, L-gates
- Planet deposits and permanent planet/system modifiers
- Enclaves and leviathans (Artisan Troupe, Curator Order, traders, Salvagers, Ether Drake, Infinity Machine)
- Ambient fauna (crystals, mining drones, amoebas, tiyanki, void clouds)
- Pre default empires, Fallen Empires, marauders, and primitives — starbases, colonies, layouts, pops, fleets, armies/transports, leaders

If L-gates were already open in the save, they start open. The Shroudwalker Coven is recreated at the save’s nexus (Pre name kept). Shroud Beacon starbases are **not** copied; build a beacon to use a copied shroud tunnel.

You pick a new empire in the creator. Continuum does **not** play you as the old save’s player. If your new empire matches a Pre empire (species class, authority, civics, ethics, non-default origin), that Pre empire is restored as **Old {name}** with reversed flag colors and **O.** leaders.

Intro title is **Continuum Present**. First line: **Welcome to your galaxy...continued!**

## How to use

1. Install Python 3 and Stellaris 4.4.x.
2. Open the old save in current Stellaris and save a **local** copy (not cloud).
3. Enable Continuum. Put this `continuum/` folder in your Stellaris `mod` directory (Workshop subscribers already have it).
4. From that folder: `python continuum_parser.py`
5. Pick the save. Wait for it to finish.
6. New Game → pick an empire → **Galaxy Size: Continuum Present** (any galaxy shape).
7. Turn extra wormholes / gateways / hyperlanes **off** if you want a near-exact import. Console `play 0` then `observe` to inspect before a real run.

`supported_version` is `v4.4.*`. Values like `v4.*` are rejected by 4.4 and can hide the mod.

## Repo vs play folder

This git repo is **source** (parser, descriptor, localisation, thumbnail). It does **not** contain a parsed galaxy.

Stellaris loads the **runtime** copy, usually:

`Documents/Paradox Interactive/Stellaris/mod/continuum`

Copy the Python modules (`continuum_parser.py`, `continuum_empires.py`, `continuum_cp.py`, `continuum_fauna.py`, `continuum_player.py`), `descriptor.mod`, `thumbnail.png`, and `localisation/` into that folder (or run the parser from a copy of this `continuum/` directory placed there). The parser **writes generated files next to itself** — map, initializers, events — from the save you pick. Do not commit those generated files (`map/`, `common/`, `events/`).

## 0.7.0

New-polity start. Pre default empires, Fallen Empires, and marauders return as living NPCs with starbases and colonies. Spawn weight prefers unowned or border systems near a matching Pre species class. Cores are not valid starts.

## 0.6.0

Unique NPCs and ambient fauna from the save (vanilla unique system inits do not run on a static map): Artisan Troupe, Curator Order, trader enclaves, Salvagers, Ether Drake, Infinity Machine, crystals, mining drones, amoebas, tiyanki, void clouds.

## 0.5.0

Planet **deposits** and permanent **planet/system modifiers** from the save.

## 0.4.0

Natural wormhole pairs, L-gates / L-cluster, Shroudwalker Coven at the save’s nexus, shroud-tunnel holes. Shroud Beacon starbases are not copied.

Stellaris 4.4.6 script reference: [OldEnt triggers / modifiers / effects](https://github.com/OldEnt/stellaris-triggers-modifiers-effects-list/tree/master).

## License

GNU Affero General Public License v3.0 — see [LICENSE](LICENSE).

© 2026 ITC Gamers
