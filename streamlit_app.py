from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
PLANNER_SCRIPT = APP_DIR / "merch_planner.py"
VISUAL_SCRIPT = APP_DIR / "planogram_visual.py"
FIXTURE_TEMPLATE = APP_DIR / "fixtures_input.csv"
ASSUMPTIONS_FILE = APP_DIR / "assumptions.json"

st.set_page_config(page_title="Merchandise Planner", page_icon="🛍️", layout="wide")
st.title("Merchandise Planner")
st.caption("Build a store planogram, estimate stock requirements, and download the results.")


def load_default_assumptions() -> str:
    try:
        return json.dumps(json.loads(ASSUMPTIONS_FILE.read_text(encoding="utf-8")), indent=2)
    except (OSError, json.JSONDecodeError) as exc:
        st.error(f"Could not load the default assumptions file: {exc}")
        st.stop()


with st.sidebar:
    st.header("Planner inputs")
    st.download_button(
        "Download fixture CSV template",
        data=FIXTURE_TEMPLATE.read_bytes(),
        file_name="fixtures_input.csv",
        mime="text/csv",
        width="stretch",
    )
    fixtures_upload = st.file_uploader(
        "Optional fixture CSV",
        type=["csv"],
        help="Use the fixture template columns and formats.",
    )
    sku_upload = st.file_uploader(
        "Optional SKU master CSV",
        type=["csv"],
        help="Expected columns include sku_code, dept, gender, subcat, colour, price, and perf_tier.",
    )
    simulation_weeks = st.slider("Lifecycle simulation (weeks)", min_value=0, max_value=52, value=6)

    with st.expander("Advanced assumptions", expanded=False):
        assumptions_text = st.text_area(
            "Edit assumptions JSON",
            value=load_default_assumptions(),
            height=360,
            help="Leave unchanged to use the defaults. The simulation-weeks control above takes precedence.",
        )

run_planner = st.button("Generate merchandise plan", type="primary", width="stretch")

if run_planner:
    try:
        assumptions = json.loads(assumptions_text)
        if not isinstance(assumptions, dict):
            raise ValueError("Assumptions must be a JSON object.")
        assumptions["simulate_weeks"] = simulation_weeks
    except (json.JSONDecodeError, ValueError) as exc:
        st.error(f"Invalid assumptions: {exc}")
        st.stop()

    if not PLANNER_SCRIPT.is_file():
        st.error(f"Planner script not found: {PLANNER_SCRIPT.name}")
        st.stop()
    if not VISUAL_SCRIPT.is_file():
        st.error(f"Planogram visualizer not found: {VISUAL_SCRIPT.name}")
        st.stop()

    with st.spinner("Generating the planogram and reports..."):
        try:
            with tempfile.TemporaryDirectory(prefix="merch_planner_") as temp_dir:
                temp_path = Path(temp_dir)
                output_dir = temp_path / "output"
                assumptions_path = temp_path / "assumptions.json"
                assumptions_path.write_text(json.dumps(assumptions), encoding="utf-8")

                command = [
                    sys.executable,
                    str(PLANNER_SCRIPT),
                    "--assumptions",
                    str(assumptions_path),
                    "--simulate-weeks",
                    str(simulation_weeks),
                    "--out",
                    str(output_dir),
                ]
                if fixtures_upload is not None:
                    fixtures_path = temp_path / "fixtures.csv"
                    fixtures_path.write_bytes(fixtures_upload.getvalue())
                    command.extend(["--fixtures", str(fixtures_path)])
                if sku_upload is not None:
                    sku_path = temp_path / "sku_master.csv"
                    sku_path.write_bytes(sku_upload.getvalue())
                    command.extend(["--sku-file", str(sku_path)])

                result = subprocess.run(
                    command,
                    cwd=APP_DIR,
                    capture_output=True,
                    text=True,
                    timeout=300,
                    check=False,
                )
                if result.returncode != 0:
                    st.error("The planner could not generate a plan.")
                    if result.stdout:
                        st.text_area("Planner output", result.stdout, height=180)
                    if result.stderr:
                        st.text_area("Error details", result.stderr, height=180)
                    st.stop()

                # The planner writes its built-in defaults to this file, so save the
                # effective user settings for the visualizer and workbook downloads.
                (output_dir / "assumptions.json").write_text(
                    json.dumps(assumptions, indent=2),
                    encoding="utf-8",
                )
                visual_result = subprocess.run(
                    [
                        sys.executable,
                        str(VISUAL_SCRIPT),
                        "--plan-dir",
                        str(output_dir),
                    ],
                    cwd=APP_DIR,
                    capture_output=True,
                    text=True,
                    timeout=300,
                    check=False,
                )
                if visual_result.returncode != 0:
                    st.error("The planogram was generated, but its visualizations could not be created.")
                    if visual_result.stdout:
                        st.text_area("Visualizer output", visual_result.stdout, height=180)
                    if visual_result.stderr:
                        st.text_area("Visualizer error details", visual_result.stderr, height=180)
                    st.stop()

                workbook_path = output_dir / "merch_plan.xlsx"
                plan_path = output_dir / "planogram.csv"
                visual_dir = output_dir / "visual"
                workbook_bytes = workbook_path.read_bytes()
                plan_csv = plan_path.read_bytes()
                floor_plan_bytes = (visual_dir / "planogram_floor.png").read_bytes()
                elevations_pdf_bytes = (visual_dir / "planogram_elevations.pdf").read_bytes()
                elevation_pages = [
                    (page_path.name, page_path.read_bytes())
                    for page_path in sorted(
                        visual_dir.glob("elevations_p*.png"),
                        key=lambda path: int(path.stem.removeprefix("elevations_p")),
                    )
                ]
                sheets = pd.read_excel(workbook_path, sheet_name=None)
                planner_output = result.stdout + "\n" + visual_result.stdout
        except subprocess.TimeoutExpired:
            st.error("Planning or visualization took longer than 5 minutes and was stopped.")
            st.stop()
        except (OSError, ValueError, pd.errors.ParserError) as exc:
            st.error(f"Could not read planner or visualization output: {exc}")
            st.stop()

    plan = sheets.get("Planogram", pd.DataFrame())
    chain = sheets.get("Chain_Rollout", pd.DataFrame())
    ranged_skus = int(plan["sku_code"].nunique()) if "sku_code" in plan else 0
    pieces_per_store = int(plan["qty_total"].sum()) if "qty_total" in plan else 0
    retail_value = float(plan["retail_value"].sum()) if "retail_value" in plan else 0.0
    chain_units = int(chain["total_units"].sum()) if "total_units" in chain else 0

    st.subheader("Plan overview")
    metric_columns = st.columns(4)
    metric_columns[0].metric("SKUs ranged", f"{ranged_skus:,}")
    metric_columns[1].metric("Pieces per store", f"{pieces_per_store:,}")
    metric_columns[2].metric("Stock value per store", f"INR {retail_value:,.0f}")
    metric_columns[3].metric("Chain opening units", f"{chain_units:,}")

    if "Segment_Summary" in sheets:
        st.subheader("Assortment by department")
        st.dataframe(sheets["Segment_Summary"], hide_index=True, width="stretch")

    st.subheader("Visual planogram")
    st.image(floor_plan_bytes, caption="Schematic store floor plan", width="stretch")
    st.download_button(
        "Download fixture elevation PDF",
        data=elevations_pdf_bytes,
        file_name="planogram_elevations.pdf",
        mime="application/pdf",
        width="stretch",
    )
    for page_name, page_bytes in elevation_pages:
        with st.expander(f"Fixture elevations — {page_name}"):
            st.image(page_bytes, caption=page_name, width="stretch")

    st.subheader("Planner reports")
    sheet_name = st.selectbox("Choose a report", list(sheets))
    st.dataframe(sheets[sheet_name], hide_index=True, width="stretch")

    download_columns = st.columns(2)
    download_columns[0].download_button(
        "Download complete Excel workbook",
        data=workbook_bytes,
        file_name="merch_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    download_columns[1].download_button(
        "Download planogram CSV",
        data=plan_csv,
        file_name="planogram.csv",
        mime="text/csv",
        width="stretch",
    )

    with st.expander("Planner run details"):
        st.code(planner_output or "Planner completed without console output.")
else:
    st.info("Choose optional input files or edit assumptions, then select **Generate merchandise plan**.")
