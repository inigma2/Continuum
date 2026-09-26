# Continuum

Stellaris 4.5.* (Cygnus) save → static galaxy. GitHub is **source**. Play/upload folder is Documents `mod\continuum`. 4.4 saves still parse; 4.5 cannot load them in-game.

## Do not regress

- Paradox Launcher uploads **only** `E:\Users\inigma\Documents\Paradox Interactive\Stellaris\mod\continuum`. Never retarget `continuum.mod` `path=` at `workshop-package`.
- `remote_file_id="3554276594"` stays on the stub and `descriptor.mod`.
- `continuum_verify.py` is GitHub-only. Do not ship it on Steam.
- Player is always a **new polity**. Date stays 2200. In-game empire creation, not a prescripted player copy.
- Spawn-weight on owned colonies is **intentional** unless the user asks to change it. Do not drop a remnant’s last system to make a start.

## QA large saves

Never load a full `gamestate` into chat. Write a short Python script (zipfile + `continuum_empires._name_from_block` / flag regex), print the subset, keep a report under `D:\AI\Continuum\_qa_*.txt`. Crisis-scale saves are tens of millions of characters.

## Names

Country names come from save name-blocks (`%ADJECTIVE%`, `AofB`). Use species loc (`SPEC_X`), not `SPEC_X_planet` (that is the homeworld). Reject `$affix$` / `$base$` templates. Successor titles are unique species+government nouns; if the noun pool is exhausted, use `… of {system}`, never `Assembly 2`.

## After Steam scrub

Re-parse the play save so `map/` / `events/` come back. Do not leave the scrubbed source-only folder as the daily play copy.
