# Continuum QA and workshop

- Inspect saves with scripts, not by reading `gamestate` into the model.
- Newest Post is newest `Post.sav` / `Test.sav` mtime under Stellaris `save games` unless the user names a file.
- Workshop scrub deletes `map/`, `common/`, `events/`, `prescripted_countries/`. Keep Python modules including `continuum_aged.py`.
- Launcher `path=` must remain Documents `mod\continuum`.
