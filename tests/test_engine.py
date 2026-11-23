"""Engine tests on synthetic data (no Enterprise data or templates in this repo)."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import engine

HEADER = "Job:LID 1 AID 1 Dig 01A Profile,Version:26.10,Units:USSurveyFeet,"


def _pass(start_counter, labels, z_shift=0.0, names=None):
    lines, c = [], start_counter
    for i, label in enumerate(labels):
        name = names[i] if names and names[i] else f"{label} {c:04d}"
        lines.append(f"{name},{8 + 0.4 * i:.3f},{-18 - 0.92 * i:.3f},{-1.9 + 0.1 * i + z_shift:.3f}")
        c += 1
    return lines


LABELS = ["2' D/S GW 181000", "1' D/S GW 181000", "D/S GW 181000"] + \
         [f"{n}' U/S GW 181000" for n in range(1, 13)]


def csv_bytes(*passes):
    rows = [HEADER, "Dig BM,0,0,0", "Dig BS,0,87.323,-2.452"]
    for p in passes:
        rows += p
    return "\n".join(rows).encode()


def test_parse_skips_control_and_reads_units():
    parsed = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS)))
    assert parsed.units == "USSurveyFeet" and not parsed.errors
    assert len(parsed.control) == 2 and len(parsed.shots) == 15


def test_metric_job_rejected():
    parsed = engine.parse_trimble_csv(b"Job:x,Units:Meters,\nA 0001,1,2,3\n")
    assert parsed.errors


def test_mislabeled_shots_get_suggestions():
    bad = ["1' D/S GW 1181000", "1' D/S GW 1181001", "D/S GW 181002"] + [None] * 12
    data = csv_bytes(_pass(1, LABELS, names=bad), _pass(16, LABELS, z_shift=0.005))
    parsed = engine.parse_trimble_csv(data)
    engine.review_names(parsed.shots)
    first = parsed.shots[:3]
    assert [s.suggestion for s in first] == ["2' D/S GW 181000 0001", "1' D/S GW 181000 0002",
                                             "D/S GW 181000 0003"]
    assert all(s.flags for s in first)
    assert not any(s.flags for s in parsed.shots[3:])


def test_name_normalisation():
    assert engine.canonical_station("32'' u/s gw 181000")[0] == "32' U/S GW 181000"
    assert engine.canonical_station("1' us gw 5000")[0] == "1' U/S GW 5000"


def test_unchanged_label_is_caught():
    labels = list(LABELS)
    labels[6] = labels[5]          # tech forgot to change the footage
    data = csv_bytes(_pass(1, LABELS), _pass(16, labels))
    parsed = engine.parse_trimble_csv(data)
    engine.review_names(parsed.shots)
    assert parsed.shots[15 + 6].flags
    assert parsed.shots[15 + 6].suggestion.startswith(LABELS[6])


def test_build_pairs_measurements_and_flags_limit():
    data = csv_bytes(_pass(1, LABELS), _pass(16, LABELS, z_shift=0.03))  # 0.36 in
    parsed = engine.parse_trimble_csv(data)
    engine.review_names(parsed.shots)
    build = engine.build_profile(parsed.shots, {s.index: s.suggestion for s in parsed.shots}, 0)
    assert not build.errors and len(build.stations) == 15
    assert [x.counter for x in build.stations[0].shots] == [1, 16]
    assert len(build.failures) == 15
    assert abs(build.stations[0].max_difference - 0.36) < 1e-6


def test_profile_split_by_counter():
    data = csv_bytes(_pass(1, LABELS), _pass(16, LABELS), _pass(1001, LABELS), _pass(1016, LABELS))
    parsed = engine.parse_trimble_csv(data)
    engine.review_names(parsed.shots)
    names = {s.index: s.suggestion for s in parsed.shots}
    assert len(engine.build_profile(parsed.shots, names, 1).stations) == 15


def test_merge_drops_repeated_shots_from_cumulative_exports():
    first = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS)))
    both = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS), _pass(16, LABELS)))
    merged = engine.merge_parsed([first, both])
    assert len(merged.shots) == 30
    assert [s.index for s in merged.shots] == list(range(30))


def test_west_positive_flips_y_only():
    parsed = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS), _pass(16, LABELS)))
    engine.review_names(parsed.shots)
    st = engine.build_profile(parsed.shots, {s.index: s.suggestion for s in parsed.shots}, 0).stations[0]
    n, e, z = st.average
    st.west_positive = True
    assert st.average == (n, -e, z)


def test_bare_weld_name_and_order_by_shot_number():
    assert engine.canonical_station("gw 3967000")[0] == "GW 3967000"
    assert engine.canonical_station("3' gw 500") is None
    # station skipped in pass 1 and shot later with its reserved number
    rows = [HEADER, "BM,0,0,0"]
    order = [1, 2, 4, 3, 5, 6, 7, 9, 8, 10]          # file order of shot numbers
    labels = {1: "3' us gw 500", 2: "gw 500", 3: "3' ds gw 500", 4: "6' ds gw 500",
              5: "9' ds gw 500"}
    dist = {1: -3, 2: 0, 3: 3, 4: 6, 5: 9}
    for c in order:
        k = c if c <= 5 else c - 5
        rows.append(f"{labels[k]} {c:04d},{10 + dist[k] + 0.003 * (c > 5):.3f},-20,-1")
    parsed = engine.parse_trimble_csv("\n".join(rows).encode())
    engine.review_names(parsed.shots)
    build = engine.build_profile(parsed.shots, {s.index: s.suggestion for s in parsed.shots}, 0)
    assert not build.errors and not build.warnings
    assert [s.station_id for s in build.stations] == [
        "3' U/S GW 500", "GW 500", "3' D/S GW 500", "6' D/S GW 500", "9' D/S GW 500"]
    assert [x.counter for x in build.stations[2].shots] == [3, 8]


def test_question_mark_coordinates_explained():
    parsed = engine.parse_trimble_csv(HEADER.encode() + b"\nbs2,?,?,?,\n3' us gw 1 0001,?,?,?,\n")
    assert parsed.errors and "backsight azimuth" in parsed.errors[0]


def test_numbered_benchmark_and_backsight_are_control():
    data = (HEADER + "\ndig 19a bm 11.19.26,0,0,0\ndig 19a bs2 11.19.26,87.3,0.2,-2.4"
            "\ndig 19a bm2 11.19.26,0,0,0\n3' us gw 1 0001,10,-20,-1\n").encode()
    parsed = engine.parse_trimble_csv(data)
    assert len(parsed.control) == 3 and len(parsed.shots) == 1


def test_one_station_off_label_is_resolved_by_geometry():
    """Pass 2 shot at the 30' mark labeled 27' (the dig 19A pattern): the 30' spot has
    one '30'' and one '27'' shot. The fix must name it 30', not split 27' across spots."""
    rows = [HEADER, "bm,0,0,0"]
    feet = [0, 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36]
    def name(f):
        return f"{f}' ds gw 500" if f else "gw 500"
    c = 1
    # pass 2 is written first in the file, so a file-order tie-break would pick the bad label
    for f in feet:
        label = name(27) if f == 30 else name(f)
        rows.append(f"{label} {c + len(feet):04d},{10 + f + 0.004:.3f},-20,-1")
        c += 1
    c = 1
    for f in feet:
        rows.append(f"{name(f)} {c:04d},{10 + f:.3f},-20,-1")
        c += 1
    parsed = engine.parse_trimble_csv("\n".join(rows).encode())
    engine.review_names(parsed.shots)
    bad = next(s for s in parsed.shots if s.counter == 11 + len(feet))
    assert bad.suggestion.startswith("30' D/S GW 500")
    build = engine.build_profile(parsed.shots, {s.index: s.suggestion for s in parsed.shots}, 0)
    assert not build.errors, build.errors
    assert all(len(st.shots) == 2 for st in build.stations)


def test_position_mode_drops_wrong_height_and_keeps_pipe_order():
    rows = [HEADER, "bm,0,0,0"]
    feet = [-3, 0, 3, 6, 9, 12]
    labels = ["3' us gw 7", "gw 7", "3' ds gw 7", "6' ds gw 7", "9' ds gw 7", "12' ds gw 7"]
    c = 1
    for p in range(2):                     # labels shifted one mark in pass 2
        for i, f in enumerate(feet):
            z = -1 - (5.0 if (p, i) == (0, 1) else 0)          # one shot at the 5 ft rod height
            lab = labels[min(i + p, len(labels) - 1)] if i >= 3 else labels[i]
            rows.append(f"{lab} {c:04d},{20 + f + 0.002 * p:.3f},-30,{z:.3f}")
            c += 1
    parsed = engine.parse_trimble_csv("\n".join(rows).encode())
    engine.review_names(parsed.shots)
    spots, gw, notes = engine.position_layout(parsed.shots, 0)
    names = engine.position_names(spots)
    assert names[1] is None                                  # wrong-height shot left out
    build = engine.build_profile(parsed.shots, names, 0, engine.pipe_order(spots))
    assert not build.errors
    assert [s.station_id for s in build.stations] == [
        "3' U/S GW 7", "GW 7", "3' D/S GW 7", "6' D/S GW 7", "9' D/S GW 7", "12' D/S GW 7"]


def test_wrong_target_height_shot_left_out():
    rows = [HEADER, "bm,0,0,0"]
    for p, start in ((0, 1), (1, 6)):
        for i in range(5):
            z = -4.0 - (0.455 if (p, i) == (0, 2) else 0)        # one shot True height 0
            rows.append(f"{3 * i}' ds gw 9 {start + i:04d},{20 + 3 * i + 0.003 * p:.3f},-30,{z:.3f}")
    parsed = engine.parse_trimble_csv("\n".join(rows).encode())
    engine.review_names(parsed.shots)
    build = engine.build_profile(parsed.shots, {s.index: s.suggestion for s in parsed.shots}, 0)
    st = build.stations[2]
    assert [x.counter for x in st.shots] == [8]
    assert any("0003 left out" in w for w in build.warnings)


def test_one_file_per_profile_numbered_from_0001():
    p0 = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS), _pass(16, LABELS)))
    p1 = engine.parse_trimble_csv(
        b"Job:AID 1 Dig 01A Profile 1,Version:26.10,Units:USSurveyFeet,\n" +
        csv_bytes(_pass(1, LABELS, z_shift=-0.04), _pass(16, LABELS, z_shift=-0.04)).split(b"\n", 1)[1])
    assert engine.detect_file_profile(p0, "dig 1a.csv") is None
    assert engine.detect_file_profile(p1, "x.csv") == 1
    assert engine.detect_file_profile(p0, "Dig 1A profile 2.csv") == 2
    merged = engine.merge_parsed([engine.apply_file_profile(p0, None),
                                  engine.apply_file_profile(p1, 1)])
    engine.review_names(merged.shots)
    assert not any(s.flags for s in merged.shots)          # repeated numbers across profiles are fine
    names = {s.index: s.suggestion for s in merged.shots}
    b0 = engine.build_profile(merged.shots, names, 0)
    b1 = engine.build_profile(merged.shots, names, 1)
    assert len(b0.stations) == len(b1.stations) == 15
    assert abs(b1.stations[0].average[2] - b0.stations[0].average[2] + 0.04) < 1e-9


def test_two_collectors_same_profile():
    first = engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS[:8]), _pass(9, LABELS[:8])))
    rows = [HEADER, "bm,0,0,0"]
    for p, start in ((0, 1), (1, 10)):
        for i, lab in enumerate(LABELS[6:]):
            k = i + 6
            rows.append(f"{lab} {start + i:04d},{8 + 0.4 * k + 0.002 * p:.3f},{-18 - 0.92 * k:.3f},{-1.9 + 0.1 * k:.3f}")
    second = engine.parse_trimble_csv("\n".join(rows).encode())
    merged = engine.merge_parsed([first, second])
    engine.review_names(merged.shots)
    assert not any(s.flags for s in merged.shots)       # both start at 0001: not a problem
    names = {s.index: s.suggestion for s in merged.shots}
    build = engine.build_profile(merged.shots, names, 0)
    assert len(build.stations) == 15 and not build.errors
    assert "share a coordinate frame" in engine.cross_file_check(merged.shots, names, 0)[0]
    # second crew in another frame: shifted 2 ft
    shifted = "\n".join(r if not r[0].isdigit() else
                        ",".join([r.split(",")[0], str(float(r.split(",")[1]) + 2)] + r.split(",")[2:])
                        for r in rows).encode()
    merged = engine.merge_parsed([engine.parse_trimble_csv(csv_bytes(_pass(1, LABELS[:8]), _pass(9, LABELS[:8]))),
                                  engine.parse_trimble_csv(shifted)])
    engine.review_names(merged.shots)
    names = {s.index: s.suggestion for s in merged.shots}
    assert "different coordinate frames" in engine.cross_file_check(merged.shots, names, 0)[0]
