from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def test_documented_pilot_acceptance_command_passes(tmp_path: Path):
    output_dir = tmp_path / "pilot-acceptance"
    completed = subprocess.run(
        [
            sys.executable,
            "scripts/run_pilot_acceptance.py",
            "--output-dir",
            str(output_dir),
        ],
        cwd=Path(__file__).parents[1],
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert "ACCEPTANCE=PASS" in completed.stdout
    assert "disputed_blank_analytics=PASS" in completed.stdout
    assert (output_dir / "input_manifest.json").is_file()
    assert (output_dir / "run_control.xlsx").is_file()
    assert sorted(path.name for path in (output_dir / "export").glob("*.xlsx")) == [
        "2024-12-31__АТ.xlsx",
        "2024-12-31__ГК.xlsx",
    ]
