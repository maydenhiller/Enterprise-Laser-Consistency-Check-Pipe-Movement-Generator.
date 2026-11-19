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
