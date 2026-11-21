"""Smoke test: run the reader on the example workbook and check key results.

    python tests/smoke_test.py
"""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.join(HERE, "..", "excel-timing-chart-reader")
READER = os.path.join(SKILL, "scripts", "xlsx_timing_reader.py")
BOOK = os.path.join(SKILL, "examples", "timing_styles.xlsx")


def run(*args):
    r = subprocess.run([sys.executable, READER, *args], capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f"FAILED: {' '.join(args)}\n{r.stderr}")
    return r.stdout


def timing(out, sheet):
    run("timing", BOOK, "--sheet", sheet, "--region", "R1", "--out", out, "--no-plot")
    tag = sheet.replace(" ", "_")
    with open(os.path.join(out, f"timing_{tag}_R1.json"), encoding="utf-8") as f:
        return json.load(f)


def main():
    out = tempfile.mkdtemp()
    run("inventory", BOOK)
    scene = run("scene", BOOK, "--out", out, "--no-plot")
    assert "legend (colour key)" in scene, "legend not separated from the chart"

    bars = timing(out, "Bars few arrows")
    assert [l["name"] for l in bars["lanes"]] == ["REQ", "GNT", "WR_DATA", "ACK", "IRQ"]
    assert len(bars["arrows"]) == 1 and bars["inferred"], "missing arrows should be inferred"

    wave = timing(out, "Waveforms")
    assert "CLK" in wave["clocks"] and abs(wave["clocks"]["CLK"]["period"] - 20) < 1e-6
    tsu = [m for m in wave["measured"] if m["label"] == "tSU"]
    assert tsu and abs(abs(tsu[0]["delta"]) - 10) < 1e-6, "tSU should measure 10"
    assert wave["handshakes"], "REQ/ACK handshake not found"

    borders = timing(out, "Borders")
    data = [l for l in borders["lanes"] if l["name"] == "DATA"][0]
    assert [s["label"] for s in data["spans"]] == ["00", "A5", "00"], "bus drawn with borders misread"

    cells = timing(out, "Cells")
    assert {l["name"] for l in cells["lanes"]} == {"EN", "BUSY", "DONE"}

    nolab = timing(out, "No labels")
    assert all(l["name_source"].startswith("UNNAMED") for l in nolab["lanes"])
    print("all checks passed")


if __name__ == "__main__":
    main()
