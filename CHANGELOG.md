# Changelog

## 0.8.5 — 2026-09-26

Stellaris **4.5.*** (Cygnus). Continuum Present and Continuum Aged.

### 4.5 Cygnus
- `supported_version` is `v4.5.*`.
- Reads Cygnus country `ethos = { ethics = { "ethic_…" } }` (4.4 `ethic="…"` still works).
- `create_pop_group` pins ethos so 4.5 does not randomize colony ethics.
- Machine remnants emit `class = MACHINE` and `trait_machine_unit` (4.5 rejects `auth_machine_intelligence` on ROBOT).
- Skip organic robot-pop techs on machine countries (robotic/droid/synthetic workers, robomodding, subdermal stimulation).

### Continuum Aged — Fallen
- Largest living remnant becomes a real `fallen_empire`: 3–5 system Gaia core, FE civics/tech, citadel without a non-designable `NAME_FE_Starbase` on `create_starbase`.
- Optional second Fallen if there were many Pre defaults and the runner-up still holds over 10% of systems. Cap 4 new Fallens (Pre FEs extra).
- Hive/machine remnants use gestalt Fallen civics. Surplus systems go to a neighbor under the 20% cap or fallow.
- Fallow (24–34%) is applied **after** Fallen shrink so the band is not overshot.
- Living remnant flag indices match event targets. Extinct capitals are not tagged as owners.

### Unclaimed megastructures
- Aged empty space: ruin if vanilla has a ruined/destroyed type (relays, gateways, arc furnaces, catapults, rings). Skip if there is no ruined type (Dyson swarms, Grand Archives, empty habitat orbitals).
- L-gates stay live. Habitats with pops stay. Pre-FTLs do not keep industrial megas.

### Sealed clusters
- L-cluster, `chosen_system`, and Formless `azilash` stay off the galactic hyperlane net (vanilla: L-gates / strange wormhole only).
- Density stitching cannot leave orphan islands, and cannot bridge those sealed systems.
- Vanilla `ancrel.12050` will not spawn a second Chosen cluster if `chosen_system` or `lcluster1` already exists.
- Formless `azilash` is not a player start. English loc remains lowercase (`azilash`) as Paradox wrote it.

### Primitives
- Leftover Pre-FTLs that reach FTL get successor names (Directorate, Senate, Autocracy, Hive nouns), not “Civilization”. Remaining primitives keep their Pre names. New pre-FTLs still use vanilla `name = random`.

### Not in 0.8.5
- Graves / archaeology (0.8.6).
- Merged stars, age slider, in-game parser.
- Wars, live Cosmic Storms objects, federation laws (membership `join_alliance` only).
- Nomad arkship restore (you may found a nomad in the empire creator).
