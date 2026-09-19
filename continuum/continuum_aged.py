"""Continuum Aged: parse-time clock advance (orbits, galactic drift, hyperlanes)."""
import copy
import math
import random
from collections import defaultdict

AGED_YEARS = 10000
INNER_FRAC = 0.33
OUTER_FRAC = 0.66


def deepcopy_galaxy(galaxy_data, nebulas=None):
    return copy.deepcopy(galaxy_data), copy.deepcopy(nebulas or [])


def _kepler_delta_deg(dist, years):
    a = max(float(dist), 1.0)
    period = max(0.15, (a / 25.0) ** 1.5)
    return (float(years) / period) * 360.0


def advance_system_orbits(hierarchy_root, years=AGED_YEARS):
    if not hierarchy_root:
        return

    def xy(body):
        return float(body.get("abs_x") or 0), float(body.get("abs_y") or 0)

    def walk(parent, orig_parent_xy):
        opx, opy = orig_parent_xy
        npx, npy = xy(parent)
        for child in parent.get("children") or []:
            ox, oy = xy(child)
            dx, dy = ox - opx, oy - opy
            dist = math.hypot(dx, dy)
            rad = math.radians(_kepler_delta_deg(dist, years))
            c, s = math.cos(rad), math.sin(rad)
            child["abs_x"] = npx + dx * c - dy * s
            child["abs_y"] = npy + dx * s + dy * c
            walk(child, (ox, oy))

    walk(hierarchy_root, xy(hierarchy_root))


def advance_orbits(galaxy_data, years=AGED_YEARS):
    for sys in galaxy_data:
        advance_system_orbits(sys.get("hierarchy_root"), years)


def _smoothstep(t):
    t = min(max(t, 0.0), 1.0)
    return t * t * (3.0 - 2.0 * t)


def spin_degrees(frac):
    """Inner 145–210°, mid 30–60°, outer 10–33°. Smooth blend, no hard rings."""
    f = min(max(frac, 0.0), 1.0)
    knots = (
        (0.00, 208.0),
        (0.30, 152.0),
        (0.37, 58.0),
        (0.62, 38.0),
        (0.82, 22.0),
        (1.00, 11.0),
    )
    for i in range(len(knots) - 1):
        f0, a0 = knots[i]
        f1, a1 = knots[i + 1]
        if f <= f1:
            t = _smoothstep((f - f0) / (f1 - f0) if f1 > f0 else 1.0)
            return a0 + (a1 - a0) * t
    return knots[-1][1]


def _apply_spin(x, y, cx, cy, rmax, years=None):
    dx, dy = x - cx, y - cy
    r = math.hypot(dx, dy)
    frac = (r / rmax) if rmax else 0.0
    # Negative angle = clockwise on the galaxy map (trailing spiral arms).
    dtheta = -math.radians(spin_degrees(frac))
    c, s = math.cos(dtheta), math.sin(dtheta)
    return cx + dx * c - dy * s, cy + dx * s + dy * c


def drift_galaxy(galaxy_data, nebulas=None, years=AGED_YEARS):
    xs = [float(s.get("x") or 0) for s in galaxy_data]
    ys = [float(s.get("y") or 0) for s in galaxy_data]
    if not xs:
        return 0.0, 0.0, 1.0
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    rmax = max(math.hypot(x - cx, y - cy) for x, y in zip(xs, ys)) or 1.0
    for s in galaxy_data:
        nx, ny = _apply_spin(float(s.get("x") or 0), float(s.get("y") or 0), cx, cy, rmax)
        s["x"] = round(nx, 3)
        s["y"] = round(ny, 3)
    for n in nebulas or []:
        nx, ny = _apply_spin(float(n.get("x") or 0), float(n.get("y") or 0), cx, cy, rmax)
        n["x"] = round(nx, 3)
        n["y"] = round(ny, 3)
    return cx, cy, rmax


def _dist(a, b):
    return math.hypot(float(a.get("x") or 0) - float(b.get("x") or 0), float(a.get("y") or 0) - float(b.get("y") or 0))


def _is_lcluster(sys):
    for fl in sys.get("flags") or []:
        s = str(fl)
        if s in ("lcluster", "lcluster1", "lcluster_lgate", "terminal_egress") or s.startswith("lcluster"):
            return True
    return False


def _is_unstable_star(sys):
    cls = str(sys.get("system_star_class") or sys.get("star_class") or "").lower()
    return any(tok in cls for tok in ("black_hole", "neutron", "pulsar", "magnetar"))


def _zone(frac):
    if frac < INNER_FRAC:
        return "inner"
    if frac < OUTER_FRAC:
        return "mid"
    return "outer"


def _xy(sys):
    return (float(sys.get("x") or 0), float(sys.get("y") or 0))


def _orient(p, q, r):
    return (q[1] - p[1]) * (r[0] - q[0]) - (q[0] - p[0]) * (r[1] - q[1])


def _segments_cross(p1, p2, p3, p4, eps=1e-9):
    if p1 == p3 or p1 == p4 or p2 == p3 or p2 == p4:
        return False
    o1, o2 = _orient(p1, p2, p3), _orient(p1, p2, p4)
    o3, o4 = _orient(p3, p4, p1), _orient(p3, p4, p2)
    if abs(o1) < eps and abs(o2) < eps and abs(o3) < eps and abs(o4) < eps:
        return False
    return (o1 > 0) != (o2 > 0) and (o3 > 0) != (o4 > 0)


def _edge_list(adj):
    out = []
    seen = set()
    for u, vs in adj.items():
        for v in vs:
            key = tuple(sorted((u, v)))
            if key in seen:
                continue
            seen.add(key)
            out.append(key)
    return out


def _crosses_any(a, b, adj, pos):
    pa, pb = pos[a], pos[b]
    for u, v in _edge_list(adj):
        if _segments_cross(pa, pb, pos[u], pos[v]):
            return True
    return False


def _add_edge(adj, a, b):
    adj[a].add(b)
    adj[b].add(a)


def _del_edge(adj, a, b):
    adj[a].discard(b)
    adj[b].discard(a)


def _snapshot_lanes(galaxy_data):
    by_id = {str(s.get("id")): s for s in galaxy_data}
    pairs = []
    seen = set()
    for s in galaxy_data:
        sid = str(s.get("id"))
        for t in s.get("hyperlanes") or []:
            tid = str(t)
            if tid not in by_id:
                continue
            key = tuple(sorted((sid, tid)))
            if key in seen:
                continue
            seen.add(key)
            pairs.append((sid, tid, _dist(s, by_id[tid])))
    return pairs


def _components(adj, nodes):
    seen = set()
    comps = []
    node_set = set(nodes)
    for start in nodes:
        if start in seen:
            continue
        stack = [start]
        seen.add(start)
        comp = []
        while stack:
            u = stack.pop()
            comp.append(u)
            for v in adj.get(u, ()):
                if v in node_set and v not in seen:
                    seen.add(v)
                    stack.append(v)
        comps.append(comp)
    comps.sort(key=len, reverse=True)
    return comps


def _k_nearest(sid, ids, pos, k=3, exclude=None):
    exclude = exclude or set()
    px, py = pos[sid]
    ranked = []
    for oid in ids:
        if oid == sid or oid in exclude:
            continue
        ox, oy = pos[oid]
        ranked.append((math.hypot(px - ox, py - oy), oid))
    ranked.sort()
    return [oid for _d, oid in ranked[:k]]


def planarize(adj, pos):
    dropped = 0
    for _round in range(24):
        edges = _edge_list(adj)
        drop = set()
        for i, (a, b) in enumerate(edges):
            if (a, b) in drop:
                continue
            pa, pb = pos[a], pos[b]
            la = math.hypot(pa[0] - pb[0], pa[1] - pb[1])
            for c, d in edges[i + 1 :]:
                if (c, d) in drop:
                    continue
                if not _segments_cross(pa, pb, pos[c], pos[d]):
                    continue
                lb = math.hypot(pos[c][0] - pos[d][0], pos[c][1] - pos[d][1])
                drop.add((a, b) if la >= lb else (c, d))
                break
        if not drop:
            break
        for a, b in drop:
            _del_edge(adj, a, b)
            dropped += 1
    return dropped


def rebuild_hyperlanes(galaxy_data, orig_pairs, cx, cy, rmax):
    """Vanilla-like planar neighborhood: inner rebuilt, mid mixed, outer mostly Pre."""
    by_id = {str(s.get("id")): s for s in galaxy_data}
    pos = {sid: _xy(s) for sid, s in by_id.items()}
    frac = {
        sid: (math.hypot(p[0] - cx, p[1] - cy) / rmax if rmax else 0.0)
        for sid, p in pos.items()
    }
    galactic = [sid for sid, s in by_id.items() if not _is_lcluster(s)]
    lcluster = [sid for sid, s in by_id.items() if _is_lcluster(s)]
    adj = defaultdict(set)
    kept, added, dropped = [], [], []
    new_set = set()
    pre_set = set()
    for a, b, _o in orig_pairs:
        pre_set.add(tuple(sorted((a, b))))

    orig_lens = sorted(d for _a, _b, d in orig_pairs) or [30.0]
    median_pre = orig_lens[len(orig_lens) // 2]
    max_new = max(48.0, median_pre * 1.85)

    def try_add(a, b, kind, force=False):
        if a == b or a not in by_id or b not in by_id:
            return False
        if a in adj[b]:
            return False
        if _is_lcluster(by_id[a]) != _is_lcluster(by_id[b]):
            return False
        if (
            not force
            and kind == "new"
            and _dist(by_id[a], by_id[b]) > max_new
            and (_zone(frac[a]) == "outer" or _zone(frac[b]) == "outer")
        ):
            return False
        if _crosses_any(a, b, adj, pos):
            return False
        _add_edge(adj, a, b)
        if kind == "keep":
            kept.append((a, b))
        else:
            added.append((a, b, round(_dist(by_id[a], by_id[b]), 2)))
            new_set.add(tuple(sorted((a, b))))
        return True

    # Outer + mid: keep Pre if still a nearby neighbor.
    for a, b, _orig in orig_pairs:
        if a not in by_id or b not in by_id:
            continue
        if _is_lcluster(by_id[a]) or _is_lcluster(by_id[b]):
            if tuple(sorted((a, b))) in pre_set and _is_lcluster(by_id[a]) and _is_lcluster(by_id[b]):
                try_add(a, b, "keep")
            continue
        za, zb = _zone(frac[a]), _zone(frac[b])
        near_a = set(_k_nearest(a, galactic, pos, k=4))
        near_b = set(_k_nearest(b, galactic, pos, k=4))
        still_near = b in near_a or a in near_b
        # Inner: keep Pre if it is still a short local link; don't dump all inner Pre.
        if za == "inner" and zb == "inner":
            dd = _dist(by_id[a], by_id[b])
            if still_near and dd <= max_new:
                if not try_add(a, b, "keep"):
                    dropped.append((a, b, "cross-or-dup"))
            else:
                dropped.append((a, b, "inner-stretched"))
            continue
        if za == "outer" or zb == "outer":
            if still_near:
                if not try_add(a, b, "keep"):
                    dropped.append((a, b, "cross-or-dup"))
            else:
                dropped.append((a, b, "not-near"))
            continue
        # mid (or mid–inner)
        if still_near:
            if not try_add(a, b, "keep"):
                dropped.append((a, b, "cross-or-dup"))
        else:
            dropped.append((a, b, "not-near"))

    # Fill sparse systems only while under Pre lane count. Prefer old leftover Pre.
    pre_target = len(orig_pairs)
    for sid in galactic:
        z = _zone(frac[sid])
        want = 2 if z == "outer" else 3
        if len(adj[sid]) >= want:
            continue
        pool = _k_nearest(sid, galactic, pos, k=want + 4)
        stable_first = [oid for oid in pool if not _is_unstable_star(by_id[oid])] + [
            oid for oid in pool if _is_unstable_star(by_id[oid])
        ]
        for oid in stable_first:
            if len(adj[sid]) >= want:
                break
            if _n_lanes(adj) >= pre_target:
                break
            try_add(sid, oid, "new")

    planarize(adj, pos)

    # One net except L-cluster.
    forbidden = set()
    for _ in range(len(galactic) + 4):
        comps = _components(adj, galactic)
        if len(comps) <= 1:
            break
        giant = set(comps[0])
        best = None
        for comp in comps[1:]:
            for sid in comp:
                for oid in _k_nearest(sid, list(giant), pos, k=12):
                    key = tuple(sorted((sid, oid)))
                    if key in forbidden:
                        continue
                    if _is_unstable_star(by_id[oid]):
                        continue
                    if _crosses_any(sid, oid, adj, pos):
                        continue
                    dd = _dist(by_id[sid], by_id[oid])
                    if best is None or dd < best[0]:
                        best = (dd, sid, oid)
        if best is None:
            for comp in comps[1:]:
                for sid in comp:
                    for oid in giant:
                        key = tuple(sorted((sid, oid)))
                        if key in forbidden:
                            continue
                        if _crosses_any(sid, oid, adj, pos):
                            continue
                        dd = _dist(by_id[sid], by_id[oid])
                        if best is None or dd < best[0]:
                            best = (dd, sid, oid)
        if not best:
            break
        if _n_lanes(adj) >= pre_target:
            break
        if not try_add(best[1], best[2], "new") and not try_add(best[1], best[2], "new", force=True):
            forbidden.add(tuple(sorted((best[1], best[2]))))

    _trim_to_pre_density(adj, pos, new_set, target=pre_target)

    # Keep L-cluster internal Pre lanes already added; don't stitch cluster to galaxy.

    cut = []
    for sid, s in by_id.items():
        old_n = {a if b == sid else b for a, b, _o in orig_pairs if sid in (a, b)}
        new_n = set(adj.get(sid, ()))
        if old_n and not (old_n & new_n):
            s["aged_pre_links_lost"] = True
            cut.append(sid)

    return adj, {
        "kept": len(kept),
        "dropped": len(dropped),
        "added": len(added),
        "cut": len(cut),
    }, kept, dropped, added


def _n_lanes(adj):
    return sum(len(vs) for vs in adj.values()) // 2


def _still_connected(adj, a, b):
    seen = set()
    stack = [a]
    while stack:
        u = stack.pop()
        if u == b:
            return True
        if u in seen:
            continue
        seen.add(u)
        stack.extend(adj[u])
    return False


def _trim_to_pre_density(adj, pos, new_set, target):
    """Drop longest new non-bridge lanes until count is at or below Pre."""
    ceiling = int(target)
    while _n_lanes(adj) > ceiling and new_set:
        cands = []
        for a, b in list(new_set):
            if b not in adj.get(a, ()):
                new_set.discard((a, b))
                continue
            if len(adj[a]) <= 2 or len(adj[b]) <= 2:
                continue
            d = math.hypot(pos[a][0] - pos[b][0], pos[a][1] - pos[b][1])
            cands.append((d, a, b))
        if not cands:
            break
        cands.sort(reverse=True)
        trimmed = False
        for _d, a, b in cands:
            _del_edge(adj, a, b)
            if _still_connected(adj, a, b):
                new_set.discard(tuple(sorted((a, b))))
                trimmed = True
                break
            _add_edge(adj, a, b)
        if not trimmed:
            break


def apply_adj(galaxy_data, adj):
    for s in galaxy_data:
        sid = str(s.get("id"))
        s["hyperlanes"] = sorted(adj[sid], key=lambda z: int(z) if str(z).isdigit() else 0)


def _nearest_stable(sid, by_id, used=None):
    used = used or set()
    s = by_id[sid]
    best, bestd = None, None
    for oid, o in by_id.items():
        if oid == sid or oid in used or _is_lcluster(o) or _is_unstable_star(o):
            continue
        dd = _dist(s, o)
        if bestd is None or dd < bestd:
            best, bestd = oid, dd
    return best, bestd


def convert_rim_wormholes(galaxy_data, adj, existing_pairs=None, max_n=2):
    """At most 2 remaining rim isolates: wormhole to nearest stable star."""
    by_id = {str(s.get("id")): s for s in galaxy_data}
    occupied = set()
    for pair in existing_pairs or []:
        if len(pair) >= 2:
            occupied.add(str(pair[0]))
            occupied.add(str(pair[1]))
    xs = [float(s.get("x") or 0) for s in galaxy_data]
    ys = [float(s.get("y") or 0) for s in galaxy_data]
    if not xs:
        return [], []
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    rmax = max(math.hypot(x - cx, y - cy) for x, y in zip(xs, ys)) or 1.0
    cands = []
    for sid, s in by_id.items():
        if _is_lcluster(s) or sid in occupied:
            continue
        r = math.hypot(float(s.get("x") or 0) - cx, float(s.get("y") or 0) - cy)
        if r < 0.70 * rmax:
            continue
        if adj.get(sid):
            continue
        cands.append((r, sid))
    cands.sort(reverse=True)
    out, labels = [], []
    used = set(occupied)
    names = {i: s.get("name") or i for i, s in by_id.items()}
    for _r, rim in cands:
        if len(out) >= max_n:
            break
        if rim in used:
            continue
        stable, sdist = _nearest_stable(rim, by_id, used)
        if not stable:
            continue
        out.append((rim, stable))
        used.add(rim)
        used.add(stable)
        labels.append(f"{names.get(rim, rim)} wormhole to {names.get(stable, stable)} (dist {sdist:.1f})")
    return out, labels


def _lab_pair(galaxy_data, a, b, extra=""):
    names = {str(s.get("id")): s.get("name") or str(s.get("id")) for s in galaxy_data}
    t = f"{names.get(a, a)}-{names.get(b, b)}"
    return f"{t} {extra}".strip() if extra else t


def drift_band_stats(galaxy_data, cx, cy, rmax):
    buckets = {"inner": [], "mid": [], "outer": []}
    for s in galaxy_data:
        x, y = float(s.get("x") or 0), float(s.get("y") or 0)
        frac = math.hypot(x - cx, y - cy) / rmax if rmax else 0.0
        buckets[_zone(frac)].append(spin_degrees(frac))
    out = {}
    for z, vals in buckets.items():
        if not vals:
            out[z] = "n=0"
            continue
        out[z] = f"n={len(vals)} {min(vals):.0f}–{max(vals):.0f}° mean {sum(vals)/len(vals):.0f}°"
    return out


def age_galaxy(galaxy_data, nebulas=None, years=AGED_YEARS, existing_wormholes=None):
    aged, aged_neb = deepcopy_galaxy(galaxy_data, nebulas)
    orig_pairs = _snapshot_lanes(aged)
    advance_orbits(aged, years)
    cx, cy, rmax = drift_galaxy(aged, aged_neb, years)
    adj, counts, kept, dropped, added = rebuild_hyperlanes(aged, orig_pairs, cx, cy, rmax)
    holes, hole_labels = convert_rim_wormholes(aged, adj, existing_wormholes, max_n=2)
    apply_adj(aged, adj)
    bands = drift_band_stats(aged, cx, cy, rmax)
    return aged, aged_neb, {
        "kept": counts["kept"],
        "dropped": counts["dropped"],
        "added": counts["added"],
        "cut": counts.get("cut") or 0,
        "dropped_lanes": [_lab_pair(aged, a, b, why) for a, b, why in dropped[:12]],
        "added_lanes": [_lab_pair(aged, a, b, str(d)) for a, b, d in added[:12]],
        "wormholes": holes,
        "wormhole_labels": hole_labels,
        "bands": bands,
        "max_keep": 0,
    }


HABITAT_CLASSES = frozenset({
    "pc_habitat", "pc_ringworld_habitable", "pc_cybrex", "pc_city",
    "pc_relic", "pc_cosmogenesis_world", "pc_nanotech",
})


def _galaxy_adj(galaxy_data):
    adj = defaultdict(set)
    for s in galaxy_data:
        sid = str(s.get("id"))
        for t in s.get("hyperlanes") or []:
            tid = str(t)
            adj[sid].add(tid)
            adj[tid].add(sid)
    return adj


def _is_habitat_class(cls):
    s = str(cls or "").lower()
    return s in HABITAT_CLASSES or "habitat" in s or "ringworld_habitable" in s


def _system_habitats(galaxy_data):
    hab = set()
    for sys in galaxy_data:
        sid = str(sys.get("id"))
        q = [sys.get("hierarchy_root")]
        while q:
            b = q.pop()
            if not b:
                continue
            if _is_habitat_class(b.get("planet_class") or b.get("class")):
                hab.add(sid)
            q.extend(b.get("children") or [])
    return hab


def _planet_classes(galaxy_data):
    out = {}
    for sys in galaxy_data:
        q = [sys.get("hierarchy_root")]
        while q:
            b = q.pop()
            if not b:
                continue
            pid = b.get("id")
            if pid is not None:
                out[str(pid)] = str(b.get("planet_class") or b.get("class") or "")
            q.extend(b.get("children") or [])
    return out


def _pops_in_system(emp, sid, p2s):
    n_col, pops = 0, 0
    for pid, n in (emp.get("colony_pop") or {}).items():
        if str(p2s.get(str(pid))) == str(sid):
            n_col += 1
            try:
                pops += int(float(n or 0))
            except (TypeError, ValueError):
                pops += 0
    return n_col, pops


def _comp_habitable(comp, emp, p2s, habitats):
    for sid in comp:
        n_col, _p = _pops_in_system(emp, sid, p2s)
        if n_col:
            return True
        if sid in habitats:
            return True
    return False


def _habitat_only(comp, emp, p2s, habitats, pclass=None):
    """True if this blob's colonies are all habitats (or habitats with no planetary colonies)."""
    pclass = pclass or {}
    cset = {str(s) for s in comp}
    n_hab_col = 0
    n_planet_col = 0
    for pid in (emp.get("colony_pop") or {}):
        if str(p2s.get(str(pid))) not in cset:
            continue
        if _is_habitat_class(pclass.get(str(pid))):
            n_hab_col += 1
        else:
            n_planet_col += 1
    if n_hab_col or n_planet_col:
        return n_planet_col == 0 and n_hab_col > 0
    return any(str(s) in habitats for s in comp)


def _hops_from(adj, src):
    if not src:
        return {}
    dist = {str(src): 0}
    q = [str(src)]
    i = 0
    while i < len(q):
        u = q[i]
        i += 1
        for v in adj.get(u, ()):
            v = str(v)
            if v not in dist:
                dist[v] = dist[u] + 1
                q.append(v)
    return dist


def _tiny_outpost(comp, emp, p2s):
    n_col, pops = 0, 0
    for sid in comp:
        c, p = _pops_in_system(emp, sid, p2s)
        n_col += c
        pops += p
    return n_col <= 1 and pops < 200


def _filter_emp_to_systems(emp, comp, p2s, c2p=None):
    spl = copy.deepcopy(emp)
    cset = {str(s) for s in comp}
    spl["_systems"] = [str(s) for s in comp]
    c2p = c2p or {}
    pops = {}
    names = {}
    for pid, n in (emp.get("colony_pop") or {}).items():
        if str(p2s.get(str(pid))) in cset:
            pops[str(pid)] = n
            nm = (emp.get("colony_names") or {}).get(str(pid))
            if nm:
                names[str(pid)] = nm
    spl["colony_pop"] = pops
    spl["colony_names"] = names
    owned = []
    for tok in emp.get("owned_planets") or []:
        pid = str(c2p.get(str(tok), tok))
        if str(p2s.get(pid)) in cset:
            owned.append(str(tok))
    spl["owned_planets"] = owned
    spl["starbases"] = {str(s): sz for s, sz in (emp.get("starbases") or {}).items() if str(s) in cset}
    sf = {}
    for sid, fl in (emp.get("system_fleets") or {}).items():
        if str(sid) in cset:
            sf[str(sid)] = fl
    spl["system_fleets"] = sf
    spl["dsc"] = {str(s): sz for s, sz in (emp.get("dsc") or {}).items() if str(s) in cset}
    spl["mega_owned"] = [m for m in (emp.get("mega_owned") or []) if str(m.get("sys")) in cset]
    spl["mine_planets"] = [p for p in (emp.get("mine_planets") or []) if str(p2s.get(str(p))) in cset]
    spl["research_planets"] = [p for p in (emp.get("research_planets") or []) if str(p2s.get(str(p))) in cset]
    spl["obs_planets"] = [p for p in (emp.get("obs_planets") or []) if str(p2s.get(str(p))) in cset]
    spl["named_military"] = [fl for fl in (emp.get("named_military") or []) if str(fl.get("sys")) in cset]
    spl["embarked_fleets"] = [fl for fl in (emp.get("embarked_fleets") or []) if str(fl.get("sys")) in cset]
    spl["civilian_fleets"] = [sh for sh in (emp.get("civilian_fleets") or []) if str(sh.get("sys")) in cset]
    spl["defense_platforms"] = {str(s): v for s, v in (emp.get("defense_platforms") or {}).items() if str(s) in cset}
    if pops:
        spl["capital"] = next(iter(pops))
    if spl.get("type") != "fallen_empire" or spl.get("_from_fe"):
        import continuum_empires as ce
        spl["starbases"] = ce.clamp_default_starbases(spl.get("starbases") or {})
    return spl


def _emp_score(emp, syss):
    pops = 0
    for v in (emp.get("colony_pop") or {}).values():
        try:
            pops += int(v)
        except (TypeError, ValueError):
            pass
    return (pops, len(emp.get("colony_pop") or {}), len(syss or []))


def _weed_defaults(defaults, bled, rng):
    """Keep top third, preserve middle third, extinct bottom third. 50/50 pops:systems."""
    rows = []
    for emp in defaults:
        cid = emp["id"]
        pops, _ncol, n_sys = _emp_score(emp, bled.get(cid) or [])
        rows.append((pops, n_sys, cid, emp.get("name")))
    max_p = max((r[0] for r in rows), default=1) or 1
    max_s = max((r[1] for r in rows), default=1) or 1
    ranked = sorted(
        rows,
        key=lambda r: 0.5 * (r[0] / max_p) + 0.5 * (r[1] / max_s),
        reverse=True,
    )
    n = len(ranked)
    if n == 0:
        return set(), set(), set()
    n_keep = max(1, n // 3)
    n_preserve = n // 3
    keep = {cid for _p, _s, cid, _n in ranked[:n_keep]}
    preserve = {cid for _p, _s, cid, _n in ranked[n_keep : n_keep + n_preserve]}
    extinct = {cid for _p, _s, cid, _n in ranked[n_keep + n_preserve :]}
    print(
        "Aged weed: keep "
        + ", ".join(str(n) for _p, _s, _c, n in ranked[:n_keep])
        + " | preserve "
        + ", ".join(str(n) for _p, _s, _c, n in ranked[n_keep : n_keep + n_preserve])
        + " | extinct "
        + ", ".join(str(n) for _p, _s, _c, n in ranked[n_keep + n_preserve :])
    )
    return keep, preserve, extinct


def sever_empires(plan, aged_galaxy):
    """Split Pre default/FE ownership on the Aged hyperlane graph."""
    rng = random.Random(int(AGED_YEARS) + 61)
    adj = _galaxy_adj(aged_galaxy)
    p2s = plan.get("planet_to_system") or {}
    habitats = _system_habitats(aged_galaxy)
    pclass = _planet_classes(aged_galaxy)
    c2p = plan.get("colony_to_planet") or {}
    defaults = list(plan.get("empires") or [])
    default_ids = {c["id"] for c in defaults}
    owned = {cid: [str(s) for s in syss] for cid, syss in (plan.get("owned") or {}).items()}
    capitals = {cid: str(cap) for cid, cap in (plan.get("capitals") or {}).items() if cap is not None}
    spawn = {str(s) for s in (plan.get("spawn_ids") or [])}
    names = {str(s.get("id")): s.get("name") or str(s.get("id")) for s in aged_galaxy}

    owner = {}
    for cid, syss in owned.items():
        for sid in syss:
            owner[sid] = cid
    emp_by_id = {e["id"]: e for e in defaults}

    def _sys_hab_only(sid, cid):
        emp = emp_by_id.get(cid)
        if not emp:
            return False
        return _habitat_only([sid], emp, p2s, habitats, pclass)

    # Border bleed: defaults only, 1/3 border then 1/3/hop max 3 deep.
    # Habitat-only systems stay with the parent (void or fallow later), never bleed.
    for cid in list(default_ids):
        for sid in list(owned.get(cid) or []):
            if sid in spawn or owner.get(sid) != cid or _sys_hab_only(sid, cid):
                continue
            others = {
                owner[n] for n in adj.get(sid, ())
                if owner.get(n) in default_ids and owner.get(n) != cid
            }
            if not others or rng.random() >= 1.0 / 3.0:
                continue
            target = rng.choice(sorted(others, key=str))
            owner[sid] = target
            frontier = [sid]
            depth = {sid: 0}
            while frontier:
                u = frontier.pop()
                if depth[u] >= 3:
                    continue
                for v in adj.get(u, ()):
                    if owner.get(v) != cid or v in spawn or v in depth or _sys_hab_only(v, cid):
                        continue
                    if rng.random() >= 1.0 / 3.0:
                        continue
                    depth[v] = depth[u] + 1
                    owner[v] = target
                    frontier.append(v)

    bled = defaultdict(list)
    for sid, cid in owner.items():
        bled[cid].append(sid)

    keep_ids, preserve_ids, extinct_ids = _weed_defaults(defaults, bled, rng)
    plan["aged_keep_ids"] = keep_ids
    plan["aged_preserve_ids"] = preserve_ids
    plan["aged_extinct_ids"] = extinct_ids

    fallow = set()
    keep_fallow = set()
    remnant = {}
    splinters = []
    names_by_sys = names

    for cid in extinct_ids:
        fallow.update(str(s) for s in (bled.get(cid) or []))

    for idx, emp in enumerate(defaults):
        cid = emp["id"]
        if cid in extinct_ids:
            remnant[cid] = []
            continue
        syss = bled.get(cid) or []
        comps = _components(adj, syss) if syss else []
        hab, empty = [], []
        for comp in comps:
            if _comp_habitable(comp, emp, p2s, habitats):
                hab.append(comp)
            else:
                empty.append(comp)
        for comp in empty:
            fallow.update(comp)
        cap = capitals.get(cid)
        cap_comp = next((c for c in hab if cap in c), None)
        largest = max(hab, key=len) if hab else None
        if cap_comp and largest and set(cap_comp) != set(largest):
            main = cap_comp if rng.random() < 0.5 else largest
        else:
            main = cap_comp or largest
        if main:
            remnant[cid] = list(main)
            hab = [c for c in hab if set(c) != set(main)]
        else:
            remnant[cid] = []
        # Grow remnant along Aged lanes through this empire's remaining systems
        # so we never keep owned pockets with no hyperlane to the capital net.
        grew = True
        while grew:
            grew = False
            rem = set(remnant[cid])
            for sid in syss:
                if sid in rem or sid in fallow:
                    continue
                if any(n in rem for n in adj.get(sid, ())):
                    remnant[cid].append(sid)
                    fallow.discard(sid)
                    hab = [c for c in hab if sid not in c]
                    grew = True

        ones = [c for c in hab if len(c) == 1]
        multi = [c for c in hab if len(c) > 1]
        if cid in preserve_ids:
            k = min(2, len(ones) // 8)
        else:
            k = min(3, len(ones) // 6)
        rng.shuffle(ones)
        chosen, leftover = ones[:k], ones[k:]
        for comp in leftover:
            if _tiny_outpost(comp, emp, p2s) and rng.random() < 0.05:
                chosen.append(comp)
            else:
                fallow.update(comp)
        import continuum_empires as ce
        sp = (plan.get("species") or {}).get(str(emp.get("founder_species"))) or {}
        for comp in multi + chosen:
            hab_only = _habitat_only(comp, emp, p2s, habitats, pclass)
            can_void = hab_only and not ce.is_machine(emp, sp) and "trait_nomadic" not in (sp.get("traits") or [])
            if hab_only and (not can_void or rng.random() < 0.5):
                fallow.update(comp)
                keep_fallow.update(comp)
                continue
            spl = _filter_emp_to_systems(emp, comp, p2s, c2p)
            spl["origin"] = "origin_default"
            spl["_parent_idx"] = idx
            spl["_parent_id"] = cid
            spl["_home"] = names_by_sys.get(str(comp[0]), "")
            if hab_only and can_void:
                spl["_void"] = True
                spl["origin"] = "origin_void_dwellers"
            splinters.append(spl)

    # At least one third of large kept empires split 2–3 ways (civil war / uprising).
    large_keep = [c for c in keep_ids if len(remnant.get(c) or []) >= 8]
    n_cw = max(1, (len(large_keep) + 2) // 3) if large_keep else 0
    cw_pick = list(large_keep)
    rng.shuffle(cw_pick)
    for cid in cw_pick[:n_cw]:
        emp = emp_by_id.get(cid)
        if not emp:
            continue
        idx = next((i for i, e in enumerate(defaults) if e["id"] == cid), 0)
        syss = [str(s) for s in (remnant.get(cid) or [])]
        cap = str(capitals.get(cid) or (syss[0] if syss else ""))
        hops = _hops_from(adj, cap)
        med = sorted(hops.get(s, 0) for s in syss)[len(syss) // 2] if syss else 0
        far = [s for s in syss if s != cap and hops.get(s, 0) >= max(2, med)]
        comps = [c for c in _components(adj, far) if _comp_habitable(c, emp, p2s, habitats)]
        comps.sort(key=len, reverse=True)
        n_extra = 2 if len(syss) >= 20 and rng.random() < 0.45 else 1
        taken = 0
        for comp in comps:
            if taken >= n_extra:
                break
            left = len(remnant.get(cid) or []) - len(comp)
            if left < 4:
                continue
            drop = {str(x) for x in comp}
            remnant[cid] = [x for x in remnant[cid] if str(x) not in drop]
            spl = _filter_emp_to_systems(emp, comp, p2s, c2p)
            spl["origin"] = "origin_default"
            spl["_civil_war"] = True
            spl["_parent_idx"] = idx
            spl["_parent_id"] = cid
            spl["_home"] = names_by_sys.get(str(comp[0]), "")
            splinters.append(spl)
            taken += 1
        if taken:
            print(f"Aged civil war: {emp.get('name')} -> {taken + 1} states")

    fe_remnant = {}
    fe_ftl = []
    fallens = list(plan.get("fallen") or [])
    fe_owned = plan.get("fallen_owned") or {}
    fe_caps = {cid: str(cap) for cid, cap in (plan.get("fallen_capitals") or {}).items() if cap is not None}
    for idx, fe in enumerate(fallens):
        cid = fe["id"]
        syss = [str(s) for s in (fe_owned.get(cid) or [])]
        comps = _components(adj, syss) if syss else []
        cap = fe_caps.get(cid)
        main = next((c for c in comps if cap in c), comps[0] if comps else [])
        fe_remnant[cid] = list(main or [])
        extra_hab = []
        for comp in comps:
            if set(comp) == set(main or []):
                continue
            if _comp_habitable(comp, fe, p2s, habitats):
                extra_hab.append(comp)
            else:
                fallow.update(comp)
        if extra_hab and rng.random() < 0.45:
            pick = extra_hab[0]
            extra_hab = extra_hab[1:]
            spl = _filter_emp_to_systems(fe, pick, p2s, c2p)
            spl["type"] = "default"
            spl["_parent_idx"] = idx
            spl["_from_fe"] = True
            spl["_home"] = names_by_sys.get(str(pick[0]), "")
            import continuum_empires as ce
            spl["starbases"] = ce.clamp_default_starbases(spl.get("starbases") or {})
            fe_ftl.append(spl)
        for comp in extra_hab:
            fallow.update(comp)

    def _rem_of():
        m = {}
        for cid, syss in remnant.items():
            for s in syss:
                m[str(s)] = cid
        return m

    # Interior holes only: fallow fully surrounded by one remnant.
    changed = True
    while changed:
        changed = False
        rem_of = _rem_of()
        for sid in list(fallow):
            if sid in keep_fallow or sid in spawn:
                continue
            nbs = list(adj.get(str(sid), ()))
            owners = [rem_of[str(n)] for n in nbs if str(n) in rem_of]
            foreign = [n for n in nbs if str(n) not in fallow and str(n) not in rem_of]
            if len(set(owners)) == 1 and not foreign:
                remnant.setdefault(owners[0], []).append(str(sid))
                fallow.discard(str(sid))
                changed = True
        keep_spl = []
        for spl in splinters:
            syss = [str(s) for s in (spl.get("_systems") or [])]
            parent = spl.get("_parent_id")
            rem = set(str(s) for s in (remnant.get(parent) or []))
            if (
                len(syss) == 1
                and not spl.get("_void")
                and not spl.get("_civil_war")
                and not spl.get("_buffer")
                and any(n in rem for n in adj.get(syss[0], ()))
            ):
                remnant.setdefault(parent, []).append(syss[0])
                changed = True
                continue
            keep_spl.append(spl)
        splinters = keep_spl

    # Kept empires reclaim their own cut-off systems, then 1 hop into foreign fallow.
    # Extinct cores stay empty so the player has room.
    own_former = {cid: {str(s) for s in (bled.get(cid) or [])} for cid in keep_ids}
    for _hop in range(2):
        claimed = []
        for cid in keep_ids:
            former = own_former.get(cid) or set()
            for sid in list(remnant.get(cid) or []):
                for n in adj.get(str(sid), ()):
                    n = str(n)
                    if n in spawn or n in keep_fallow or n not in fallow:
                        continue
                    if n in former or (_hop == 0 and rng.random() < 0.35):
                        claimed.append((cid, n))
        for cid, n in claimed:
            if n in fallow:
                remnant.setdefault(cid, []).append(n)
                fallow.discard(n)

    keep_spl = []
    for spl in splinters:
        syss = [str(s) for s in (spl.get("_systems") or [])]
        parent = spl.get("_parent_id")
        rem = set(str(s) for s in (remnant.get(parent) or []))
        if (
            syss
            and not spl.get("_void")
            and not spl.get("_civil_war")
            and not spl.get("_buffer")
            and any(any(str(n) in rem for n in adj.get(s, ())) for s in syss)
        ):
            remnant.setdefault(parent, []).extend(syss)
            continue
        keep_spl.append(spl)
    splinters = keep_spl

    n_sys = max(1, len(aged_galaxy))
    target_frac = rng.uniform(0.24, 0.34)
    target_n = int(round(target_frac * n_sys))

    def _shrink_cid(cid, min_keep):
        syss = [str(s) for s in (remnant.get(cid) or [])]
        if len(syss) <= min_keep:
            return False
        cap = str(capitals.get(cid) or (syss[0] if syss else ""))
        hops = _hops_from(adj, cap)
        cands = [s for s in syss if s != cap and s not in spawn]
        cands.sort(key=lambda s: hops.get(s, 99), reverse=True)
        if not cands:
            return False
        s = cands[0]
        remnant[cid] = [x for x in remnant[cid] if str(x) != s]
        fallow.add(s)
        return True

    while len(fallow) < target_n:
        shrunk = False
        for cid in list(preserve_ids):
            if _shrink_cid(cid, 2):
                shrunk = True
                break
        if shrunk:
            continue
        for cid in list(keep_ids):
            if _shrink_cid(cid, 8):
                shrunk = True
                break
        if shrunk:
            continue
        extras = [
            s for s in splinters
            if not s.get("_civil_war") and not s.get("_buffer") and len(s.get("_systems") or []) <= 2
        ]
        if extras:
            extras.sort(key=lambda s: len(s.get("_systems") or []))
            spl = extras[0]
            splinters = [s for s in splinters if s is not spl]
            fallow.update(str(x) for x in (spl.get("_systems") or []))
            continue
        break
    max_own = max(1, int(0.20 * n_sys))

    def _peel_far(syss, cap, n_drop):
        hops = _hops_from(adj, cap)
        cands = [str(s) for s in syss if str(s) != str(cap) and str(s) not in spawn]
        cands.sort(key=lambda s: hops.get(s, 99), reverse=True)
        return cands[:n_drop]

    def _make_rival(parent_emp, parent_idx, parent_cid, syss):
        if not parent_emp or not syss:
            return
        hab = [c for c in _components(adj, syss) if _comp_habitable(c, parent_emp, p2s, habitats)]
        if not hab:
            fallow.update(str(s) for s in syss)
            return
        for comp in hab:
            spl = _filter_emp_to_systems(parent_emp, comp, p2s, c2p)
            spl["origin"] = "origin_default"
            spl["_civil_war"] = True
            spl["_parent_idx"] = parent_idx
            spl["_parent_id"] = parent_cid
            spl["_home"] = names_by_sys.get(str(comp[0]), "")
            splinters.append(spl)
        leftover = set(str(s) for s in syss) - {str(s) for c in hab for s in c}
        fallow.update(leftover)

    idx_of = {e["id"]: i for i, e in enumerate(defaults)}
    for cid in list(keep_ids) + list(preserve_ids):
        syss = [str(s) for s in (remnant.get(cid) or [])]
        if len(syss) <= max_own:
            continue
        cap = str(capitals.get(cid) or (syss[0] if syss else ""))
        extra = _peel_far(syss, cap, len(syss) - max_own)
        drop = set(extra)
        remnant[cid] = [x for x in remnant[cid] if str(x) not in drop]
        if len(fallow) < target_n:
            need = target_n - len(fallow)
            to_fallow, to_rival = extra[:need], extra[need:]
            fallow.update(to_fallow)
            if to_rival:
                _make_rival(emp_by_id.get(cid), idx_of.get(cid, 0), cid, to_rival)
        else:
            _make_rival(emp_by_id.get(cid), idx_of.get(cid, 0), cid, extra)
        print(f"Aged 20% cap: {emp_by_id.get(cid, {}).get('name')} now {len(remnant.get(cid) or [])} systems")

    for spl in list(splinters):
        syss = [str(s) for s in (spl.get("_systems") or [])]
        if len(syss) <= max_own:
            continue
        cap = syss[0]
        extra = _peel_far(syss, cap, len(syss) - max_own)
        drop = set(extra)
        spl["_systems"] = [x for x in syss if x not in drop]
        parent = emp_by_id.get(spl.get("_parent_id"))
        if len(fallow) < target_n:
            need = target_n - len(fallow)
            fallow.update(extra[:need])
            rest = extra[need:]
            if rest and parent:
                _make_rival(parent, spl.get("_parent_idx") or 0, spl.get("_parent_id"), rest)
        elif parent:
            _make_rival(parent, spl.get("_parent_idx") or 0, spl.get("_parent_id"), extra)

    fat = max(1, int(0.10 * n_sys))
    buf_max = 10

    def _border_syss(syss, cap):
        owned = {str(s) for s in syss}
        hops = _hops_from(adj, cap)
        out = []
        for s in syss:
            s = str(s)
            if s == str(cap) or s in spawn:
                continue
            if any(str(n) not in owned for n in adj.get(s, ())):
                out.append(s)
        out.sort(key=lambda s: hops.get(s, 99), reverse=True)
        return out

    def _chunk_border(take):
        pieces = []
        for comp in _components(adj, take):
            cur = [str(x) for x in comp]
            while cur:
                pieces.append(cur[:buf_max])
                cur = cur[buf_max:]
        return pieces

    def _emit_buffers(parent_emp, parent_idx, parent_cid, pieces):
        n_buf = 0
        extra = pieces[5:]
        pieces = pieces[:5]
        if extra:
            fallow.update(str(s) for p in extra for s in p)
        if not parent_emp:
            fallow.update(str(s) for p in pieces for s in p)
            return 0
        for piece in pieces:
            if not piece:
                continue
            spl = _filter_emp_to_systems(parent_emp, piece, p2s, c2p)
            spl["origin"] = "origin_default"
            spl["_buffer"] = True
            spl["_parent_idx"] = parent_idx
            spl["_parent_id"] = parent_cid
            spl["_home"] = names_by_sys.get(str(piece[0]), "")
            splinters.append(spl)
            n_buf += 1
        return n_buf

    for cid in list(keep_ids) + list(preserve_ids):
        syss = [str(s) for s in (remnant.get(cid) or [])]
        if len(syss) <= fat:
            continue
        n_peel = min(max(1, len(syss) // 5), len(syss) - fat)
        cap = str(capitals.get(cid) or (syss[0] if syss else ""))
        take = _border_syss(syss, cap)[:n_peel]
        if not take:
            continue
        drop = set(take)
        remnant[cid] = [x for x in remnant[cid] if str(x) not in drop]
        n_buf = _emit_buffers(emp_by_id.get(cid), idx_of.get(cid, 0), cid, _chunk_border(take))
        print(
            f"Aged buffer: {emp_by_id.get(cid, {}).get('name')} peeled {len(take)} "
            f"-> {n_buf} fringe (1-{buf_max}) core {len(remnant.get(cid) or [])} "
            f"({100 * len(remnant.get(cid) or []) / n_sys:.1f}%)"
        )

    for spl in list(splinters):
        if spl.get("_buffer"):
            continue
        syss = [str(s) for s in (spl.get("_systems") or [])]
        if len(syss) <= fat:
            continue
        n_peel = min(max(1, len(syss) // 5), len(syss) - fat)
        take = _border_syss(syss, syss[0])[:n_peel]
        if not take:
            continue
        drop = set(take)
        spl["_systems"] = [x for x in syss if x not in drop]
        parent = emp_by_id.get(spl.get("_parent_id"))
        n_buf = _emit_buffers(parent, spl.get("_parent_idx") or 0, spl.get("_parent_id"), _chunk_border(take))
        print(
            f"Aged buffer: splinter {spl.get('_home')} peeled {len(take)} -> {n_buf} fringe"
        )

    print(
        f"Aged fallow target {target_frac:.0%} ({target_n}/{n_sys}) actual {len(fallow)} "
        f"({len(fallow) / n_sys:.0%}) splinters_left={len(splinters)} cap={max_own} fat={fat}"
    )

    prims = list(plan.get("primitives") or [])
    prim_ftl = []
    import continuum_empires as ce
    _FTL_P = {
        "stone_age": 0.15, "bronze_age": 0.22, "iron_age": 0.30,
        "late_medieval_age": 0.45, "renaissance_age": 0.55, "steam_age": 0.68,
        "industrial_age": 0.80, "machine_age": 0.85, "atomic_age": 0.90,
        "early_space_age": 0.95,
    }
    for i, p in enumerate(prims):
        sp = (plan.get("species") or {}).get(str(p.get("founder_species"))) or {}
        if ce.is_hive(p, sp) or ce.is_machine(p, sp):
            prim_ftl.append(i)
            continue
        chance = _FTL_P.get(p.get("pre_ftl_age") or "", 0.55)
        if rng.random() < chance:
            prim_ftl.append(i)
    print(f"Aged prim_ftl {len(prim_ftl)}/{len(prims)} (age-weighted)")

    occupied = set(spawn)
    for syss in remnant.values():
        occupied.update(str(s) for s in syss)
    for spl in splinters + fe_ftl:
        occupied.update(str(s) for s in (spl.get("_systems") or []))
    for p in prims:
        if p.get("system_id"):
            occupied.add(str(p["system_id"]))
    blocked = (
        "guardians_", "lcluster", "marauder", "shroud", "enclave", "lgate",
        "terminal_egress", "wenkwort", "tiyanki", "void3", "crisis",
    )
    cands = []
    for s in aged_galaxy:
        sid = str(s.get("id"))
        if sid not in fallow or sid in occupied:
            continue
        fls = [str(f) for f in (s.get("flags") or [])]
        if any(any(b in str(f) for b in blocked) for f in fls):
            continue
        hab = None
        q = [s.get("hierarchy_root")]
        while q:
            b = q.pop()
            if not b:
                continue
            pc = str(b.get("planet_class") or b.get("class") or "")
            if ce.is_habitable_class(pc) and b.get("id") is not None:
                hab = (str(b.get("id")), pc)
                break
            q.extend(b.get("children") or [])
        if hab:
            cands.append((sid, hab[0], hab[1]))
    rng.shuffle(cands)
    n_new = min(8, max(3, len(cands) // 15)) if cands else 0
    ages = ["stone_age", "stone_age", "bronze_age", "iron_age", "late_medieval_age"]
    classes = ["MAM", "REP", "AVI", "MOL", "ART", "FUN"]
    namelists = {"MAM": "MAM1", "REP": "REP1", "AVI": "AVI1", "MOL": "MOL1", "ART": "ART1", "FUN": "FUN1"}
    new_prims = []
    for sid, pid, pc in cands[:n_new]:
        cls = rng.choice(classes)
        new_prims.append({
            "system_id": sid,
            "planet_id": pid,
            "planet_class": pc,
            "age": rng.choice(ages),
            "class": cls,
            "namelist": namelists.get(cls, "MAM1"),
        })
        fallow.discard(sid)
    plan["aged_new_prims"] = new_prims
    print(f"Aged new pre-FTLs {len(new_prims)}")

    return {
        "remnant": remnant,
        "fe_remnant": fe_remnant,
        "splinters": splinters,
        "fe_ftl": fe_ftl,
        "fallow": fallow,
        "prim_ftl": prim_ftl,
        "habitats": habitats,
    }


def _strip_owner_flags(flags, prefixes):
    out = {}
    for sid, fls in (flags or {}).items():
        keep = [f for f in fls if not any(str(f).startswith(p) for p in prefixes)]
        if keep:
            out[str(sid)] = keep
    return out


def apply_aged_flags(plan, sever):
    """Rebuild Aged star/planet flags: remnant, splinters, fallow, prim FTL."""
    import continuum_empires as ce

    flags = _strip_owner_flags(plan.get("extra_flags"), ("continuum_emp_", "continuum_fe_", "continuum_prim_", "continuum_spl_", "continuum_newprim_"))
    planet_flags = _strip_owner_flags(plan.get("planet_flags"), ("continuum_emp_", "continuum_fe_", "continuum_prim_", "continuum_spl_", "continuum_newprim_"))
    c2p = plan.get("colony_to_planet") or {}
    p2s = plan.get("planet_to_system") or {}
    defaults = list(plan.get("empires") or [])
    fallens = list(plan.get("fallen") or [])
    remnant = sever["remnant"]
    fe_remnant = sever["fe_remnant"]
    rem_emps, rem_caps = [], {}
    for emp in defaults:
        syss = remnant.get(emp["id"]) or []
        e2 = _filter_emp_to_systems(emp, syss, p2s, c2p)
        e2["id"] = emp["id"]
        for sid in syss:
            e2.setdefault("starbases", {}).setdefault(str(sid), "starbase_outpost")
        rem_emps.append(e2)
        cap = (plan.get("capitals") or {}).get(emp["id"])
        rem_caps[emp["id"]] = cap if cap and str(cap) in syss else (syss[0] if syss else cap)
    fe_emps, fe_caps = [], {}
    for emp in fallens:
        syss = fe_remnant.get(emp["id"]) or []
        e2 = _filter_emp_to_systems(emp, syss, p2s, c2p)
        e2["id"] = emp["id"]
        for sid in syss:
            e2.setdefault("starbases", {}).setdefault(str(sid), "starbase_outpost")
        fe_emps.append(e2)
        cap = (plan.get("fallen_capitals") or {}).get(emp["id"])
        fe_caps[emp["id"]] = cap if cap and str(cap) in syss else (syss[0] if syss else cap)
    ce._tag_owned_planets(rem_emps, c2p, "continuum_emp", flags, planet_flags, remnant, rem_caps)
    ce._tag_owned_planets(fe_emps, c2p, "continuum_fe", flags, planet_flags, fe_remnant, fe_caps)
    spl_owned = {}
    spl_caps = {}
    for i, spl in enumerate(sever["splinters"] + sever["fe_ftl"]):
        sid0 = (spl.get("_systems") or ["0"])[0]
        oid = f"spl{i}"
        spl["id"] = oid
        spl_owned[oid] = list(spl.get("_systems") or [])
        spl_caps[oid] = sid0
        for sid in spl_owned[oid]:
            spl.setdefault("starbases", {}).setdefault(str(sid), "starbase_outpost")
    ce._tag_owned_planets(sever["splinters"] + sever["fe_ftl"], c2p, "continuum_spl", flags, planet_flags, spl_owned, spl_caps)
    for sid in sever["fallow"]:
        flags.setdefault(str(sid), []).append("continuum_aged_fallow")
    prim_ftl = set(sever.get("prim_ftl") or [])
    for idx, prim in enumerate(plan.get("primitives") or []):
        sid = prim.get("system_id")
        if not sid:
            continue
        flags.setdefault(str(sid), []).append(f"continuum_prim_{idx}")
        if idx in prim_ftl:
            flags.setdefault(str(sid), []).append(f"continuum_prim_ftl_{idx}")
        cap_col = prim.get("capital") or ((prim.get("owned_planets") or [None])[0])
        cap_pid = c2p.get(str(cap_col)) if cap_col is not None else None
        if cap_pid:
            planet_flags.setdefault(str(cap_pid), []).append(f"continuum_prim_{idx}_homeworld")
            planet_flags.setdefault(str(cap_pid), []).append(f"continuum_prim_{idx}_c0")
        for i, pid in enumerate((prim.get("colony_pop") or {}).keys()):
            planet_flags.setdefault(str(pid), []).append(f"continuum_prim_{idx}_c{i}")
    spawn = dict(plan.get("spawn_weights") or {})
    spawn_ids = {str(s) for s in (plan.get("spawn_ids") or [])}
    for idx, prim in enumerate(plan.get("primitives") or []):
        sid = prim.get("system_id")
        if sid and idx in prim_ftl:
            spawn.setdefault(str(sid), {"base": 8, "modifiers": []})
            spawn_ids.add(str(sid))
    for i, np in enumerate(plan.get("aged_new_prims") or []):
        sid = np.get("system_id")
        if sid:
            flags.setdefault(str(sid), []).append(f"continuum_newprim_{i}")
        pid = np.get("planet_id")
        if pid:
            planet_flags.setdefault(str(pid), []).append(f"continuum_newprim_{i}_homeworld")
    plan["aged_spawn_ids"] = spawn_ids
    n_spl = len(sever["splinters"])
    n_fo = len(sever["fallow"])
    n_rem = sum(len(v) for v in remnant.values())
    print(
        f"Aged severance: remnant_systems={n_rem} splinters={n_spl} "
        f"fe_ftl={len(sever['fe_ftl'])} fallow={n_fo} prim_ftl={len(sever['prim_ftl'])} "
        f"keep={len(plan.get('aged_keep_ids') or [])} "
        f"preserve={len(plan.get('aged_preserve_ids') or [])} "
        f"extinct={len(plan.get('aged_extinct_ids') or [])}"
    )
    return flags, planet_flags, spawn
