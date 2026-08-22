"""Inventory a Stellaris .sav for Continuum Pre/Post checks. Does not load the full gamestate into memory as one string."""
import io
import re
import sys
import zipfile
from collections import defaultdict

TEST_PREFIX = "TEST_"


def _brace_block(line_iterator, first_line):
    lines = [first_line]
    depth = first_line.count("{") - first_line.count("}")
    if depth <= 0 and "{" in first_line:
        return "".join(lines)
    for line in line_iterator:
        lines.append(line)
        depth += line.count("{") - line.count("}")
        if depth <= 0:
            break
    return "".join(lines)


def _name_from_block(text):
    m = re.search(r'key="([^"]+)"', text)
    if m:
        return m.group(1)
    m = re.search(r'^\s*name="([^"]+)"', text, re.MULTILINE)
    if m:
        return m.group(1)
    return None


def _read_id_block(line_iterator, header_line):
    if "{" in header_line:
        return _brace_block(line_iterator, header_line)
    open_line = next(line_iterator, "{\n")
    return header_line + _brace_block(line_iterator, open_line)


def inventory_save(path):
    systems = []
    megas = []
    bypasses = []
    hits = defaultdict(int)
    scan_keys = (
        "shroud_beacon",
        "shroud_tunnel_nexus",
        "spawned_shroud_tunnel",
        "shroud_tunnel_node",
        "lgate_base",
        "lgates_activated_globally",
        'has_star_flag = lgate',
        "lcluster1",
        "continuum_wormhole_",
    )

    with zipfile.ZipFile(path, "r") as z:
        if "gamestate" not in z.namelist():
            raise SystemExit(f"No gamestate in {path}")
        with z.open("gamestate") as raw:
            it = io.TextIOWrapper(raw, encoding="utf-8")
            section = None
            header_re = re.compile(r"^\t(\d+)=")
            for line in it:
                stripped = line.strip()
                if stripped == "galactic_object=":
                    section = "stars"
                    next(it, None)
                    continue
                if stripped == "megastructures=":
                    section = "megas"
                    next(it, None)
                    continue
                if stripped == "bypasses=":
                    section = "bypasses"
                    next(it, None)
                    continue
                if stripped in ("planets=", "natural_wormholes=", "nebula=", "starbase_instances=", "countries="):
                    if stripped != "nebula=":
                        section = stripped.rstrip("=")
                    continue

                for key in scan_keys:
                    if key in line:
                        hits[key] += 1

                if section == "stars":
                    m = header_re.match(line)
                    if m and stripped != "}":
                        block = _read_id_block(it, line)
                        sid = m.group(1)
                        name = _name_from_block(block) or f"sys_{sid}"
                        flags_m = re.search(r"flags=\s*\{([^}]*)\}", block, re.DOTALL)
                        flags = []
                        if flags_m:
                            flags = [ln.strip().split("=")[0] for ln in flags_m.group(1).splitlines() if ln.strip()]
                        systems.append({"id": sid, "name": name, "flags": flags})
                    elif stripped == "}" and line.startswith("}"):
                        section = None
                elif section == "megas":
                    m = header_re.match(line)
                    if m:
                        block = _read_id_block(it, line)
                        t = re.search(r'^\s*type="([^"]+)"', block, re.MULTILINE)
                        origin = re.search(r"origin=(\d+)", block)
                        planet = re.search(r"^\s*planet=(\d+)", block, re.MULTILINE)
                        megas.append({
                            "id": m.group(1),
                            "type": t.group(1) if t else "?",
                            "origin": origin.group(1) if origin else "?",
                            "planet": planet.group(1) if planet else "",
                        })
                    elif stripped == "}" and not line.startswith("\t"):
                        section = None
                elif section == "bypasses":
                    m = header_re.match(line)
                    if m:
                        block = _read_id_block(it, line)
                        t = re.search(r'^\s*type="([^"]+)"', block, re.MULTILINE)
                        linked = re.search(r"linked_to=(\d+)", block)
                        bypasses.append({
                            "id": m.group(1),
                            "type": t.group(1) if t else "?",
                            "linked_to": linked.group(1) if linked else "",
                        })
                    elif stripped == "}" and not line.startswith("\t"):
                        section = None

    return systems, megas, bypasses, hits


def print_inventory(path):
    systems, megas, bypasses, hits = inventory_save(path)
    print(f"=== {path} ===")
    print(f"systems={len(systems)} megastructures={len(megas)} bypasses={len(bypasses)}")
    print("\n-- TEST_* systems --")
    test = [s for s in systems if s["name"].startswith(TEST_PREFIX) or TEST_PREFIX in s["name"]]
    if not test:
        print("(none)")
    for s in test:
        flags = ",".join(s["flags"][:12]) if s["flags"] else ""
        print(f"  id={s['id']} name={s['name']} flags={flags}")
    print("\n-- megastructures --")
    by_type = defaultdict(int)
    for m in megas:
        by_type[m["type"]] += 1
        planet_note = f" planet={m['planet']}" if m["planet"] and m["planet"] != "4294967295" else ""
        if TEST_PREFIX in str(m) or m["type"] in (
            "gateway_ruined", "gateway_restored", "gateway_final", "hyper_relay",
            "dyson_sphere_ruined", "habitat_0", "lgate_base",
        ):
            print(f"  id={m['id']} type={m['type']} origin={m['origin']}{planet_note}")
    print("  counts:", dict(sorted(by_type.items(), key=lambda kv: (-kv[1], kv[0]))))
    print("\n-- bypasses --")
    by_b = defaultdict(int)
    for b in bypasses:
        by_b[b["type"]] += 1
        extra = f" linked_to={b['linked_to']}" if b["linked_to"] else ""
        if b["type"] in ("wormhole", "shroud_tunnel", "lgate", "gateway"):
            print(f"  id={b['id']} type={b['type']}{extra}")
    print("  counts:", dict(by_b))
    print("\n-- keyword hits --")
    for k, v in sorted(hits.items()):
        print(f"  {k}: {v}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python continuum_verify.py <Pre.sav> [Post.sav]")
        sys.exit(1)
    print_inventory(sys.argv[1])
    if len(sys.argv) > 2:
        print()
        print_inventory(sys.argv[2])


if __name__ == "__main__":
    main()
