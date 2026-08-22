# Continuum

A Stellaris galaxy converter — one galaxy to rule them all.

Turn a finished (or mid-game) save into a playable **static galaxy** for a new game. The Python parser reads your `.sav` and writes systems, hyperlanes, stars, planets, moons, asteroids, belts, nebulas, megastructures, and wormholes into the Continuum mod folder.

- **Current:** 0.4.0 for Stellaris **4.4.*** (Pegasus)
- Workshop: https://steamcommunity.com/sharedfiles/filedetails/?id=3554276594
- Roadmap: https://steamcommunity.com/workshop/filedetails/discussion/3554276594/596284386694022138/

## Repo vs play folder

This git repo is **source** (parser, descriptor, localisation, thumbnail). It does **not** contain a parsed galaxy.

Stellaris loads the **runtime** copy, usually:

`Documents/Paradox Interactive/Stellaris/mod/continuum`

Copy `continuum_parser.py`, `descriptor.mod`, `thumbnail.png`, and `localisation/` into that folder (or run the parser from a copy of this `continuum/` directory placed there). The parser **writes generated files next to itself** — map, initializers, events — from the save you pick. Do not commit those generated files.

## How to use

1. Install Python 3 and Stellaris 4.4.x.
2. Open the old save in current Stellaris and save a **local** copy (not cloud).
3. Put this `continuum/` folder in your Stellaris `mod` directory and enable Continuum in the launcher.
4. From that folder: `python continuum_parser.py`
5. Pick the save. Wait for it to finish.
6. New Game → pick an empire → **Galaxy Size: Continuum** (any galaxy shape).
7. Turn extra wormholes / gateways / hyperlanes off if you want a near-exact import. Console `observe` to inspect before playing.

`supported_version` is `v4.4.*`. Values like `v4.*` are rejected by 4.4 and can hide the mod.

## 0.4.0

Copies more of the save’s **structure**, not empires or starbases:

- Natural wormhole pairs (once per new game).
- L-gates and L-cluster flags. If the save had already opened the L-gates, Continuum opens them on new game.
- Shroudwalker Coven at the save’s nexus, plus shroud-tunnel holes at nexus and nodes.
- Shroud **Beacon** starbase buildings are **not** copied (a starbase always needs an owner). A new Teachers / Coven player builds a beacon to use the copied holes. Vanilla origin events still fire.

Do not commit parser output (`map/`, `common/`, `events/`).

Stellaris 4.4.6 script reference: [OldEnt triggers / modifiers / effects](https://github.com/OldEnt/stellaris-triggers-modifiers-effects-list/tree/master).

## License

GNU Affero General Public License v3.0 — see [LICENSE](LICENSE).

© 2025 ITC Gamers
