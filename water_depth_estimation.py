"""
Water Depth Estimation
======================

Python routine for estimating corrected water-column height above the seabed
from pressure-transducer records.

The script reads a pressure-transducer time series, extracts predefined
submerged intervals, estimates an air-reference offset, calculates corrected
water-column height, applies a moving average, and exports both Excel tables
and per-period figures.

Expected input columns
----------------------
- Date_Time: date and time of each pressure-transducer record
- Depth_PT_m: pressure-transducer depth reading in metres

Main outputs
------------
- corrected water-column height time series
- moving-average water-column height time series
- per-period summary statistics
- metadata and interval definitions
- raw and smoothed plots for each submerged interval
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# =============================================================================
# CONFIGURATION
# =============================================================================

INPUT_PREFIX = "INPUT_"
OUTPUT_PREFIX = "OUTPUT_"

DEFAULT_INPUT_FILE = "INPUT_PT_DATA.xlsx"
DEFAULT_INTERVAL_SET_LABEL = "September_2022"


@dataclass
class OceanographicConstants:
    """Physical settings used in the correction."""

    z_sensor: float = 0.05  # sensor height above the bed, in metres


@dataclass
class ProcessingConfig:
    """Processing settings."""

    datetime_col: str = "Date_Time"
    depth_col: str = "Depth_PT_m"
    date_format: str = "%d-%m-%Y %H:%M:%S"
    air_reference_mode: str = "previous_dry"  # global | previous_dry | surrounding_dry
    clip_negative_to_zero: bool = True
    y_axis_max_m: float = 1.2
    moving_average_window: str = "1min"


@dataclass(frozen=True)
class TimeInterval:
    """One submerged interval to process."""

    start: str
    end: str
    label: str

    def start_ts(self, fmt: str) -> pd.Timestamp:
        return pd.to_datetime(self.start, format=fmt)

    def end_ts(self, fmt: str) -> pd.Timestamp:
        return pd.to_datetime(self.end, format=fmt)


CONSTANTS = OceanographicConstants()
CONFIG = ProcessingConfig()

# Adapt this list for each campaign or deployment.
TIME_INTERVALS = [
    TimeInterval("24-09-2022 12:13:00", "24-09-2022 17:00:58", "1stPeriod"),
    TimeInterval("25-09-2022 00:50:00", "25-09-2022 04:54:58", "2ndPeriod"),
    TimeInterval("25-09-2022 12:40:00", "25-09-2022 17:40:58", "3rdPeriod"),
    TimeInterval("26-09-2022 01:06:02", "26-09-2022 05:40:58", "4thPeriod"),
    TimeInterval("26-09-2022 13:15:04", "26-09-2022 18:09:58", "5thPeriod"),
]

REQUIRED_COLUMNS = ["Date_Time", "Depth_PT_m"]

PLOT_STYLE = {
    "figsize": (12, 4.5),
    "dpi": 300,
    "line_width": 1.2,
    "grid_alpha": 0.30,
    "x_label": "Date and time",
    "y_label": "H (m)",
}


# =============================================================================
# PATHS AND INPUT
# =============================================================================

def safe_tag(text: str) -> str:
    """Return a short filesystem-safe label."""
    return str(text).strip().replace(" ", "_").replace("/", "_").replace("\\", "_")


def output_excel_path_for(input_path: Path, interval_set_label: str) -> Path:
    """Build a logical output workbook name from the input filename."""
    tag = safe_tag(interval_set_label)

    if input_path.name.startswith(INPUT_PREFIX):
        output_name = (
            input_path.name.replace(INPUT_PREFIX, OUTPUT_PREFIX, 1)
            .replace(".xlsx", f"_WATER_COLUMN_HEIGHT_{tag}.xlsx")
            .replace(".xls", f"_WATER_COLUMN_HEIGHT_{tag}.xlsx")
            .replace(".csv", f"_WATER_COLUMN_HEIGHT_{tag}.xlsx")
        )
    else:
        output_name = f"{OUTPUT_PREFIX}{input_path.stem}_WATER_COLUMN_HEIGHT_{tag}.xlsx"

    return input_path.with_name(output_name)


def output_plot_dir_for(input_path: Path, interval_set_label: str) -> Path:
    """Build a logical output directory for period plots."""
    return input_path.with_name(f"OUTPUT_PLOTS_{safe_tag(interval_set_label)}")


def read_input_data(input_file: Path, sheet: str | int | None = None) -> pd.DataFrame:
    """Read and validate an Excel or CSV pressure-transducer input file."""
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    if input_file.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(input_file, sheet_name=sheet if sheet is not None else 0)
    elif input_file.suffix.lower() == ".csv":
        df = pd.read_csv(input_file)
    else:
        raise ValueError("Input file must be .xlsx, .xls, or .csv.")

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df.copy()
    if pd.api.types.is_datetime64_any_dtype(df[CONFIG.datetime_col]):
        df[CONFIG.datetime_col] = pd.to_datetime(df[CONFIG.datetime_col])
    else:
        df[CONFIG.datetime_col] = pd.to_datetime(
            df[CONFIG.datetime_col],
            format=CONFIG.date_format,
        )
    df[CONFIG.depth_col] = pd.to_numeric(df[CONFIG.depth_col], errors="coerce")
    df = df.dropna(subset=[CONFIG.datetime_col, CONFIG.depth_col])
    df = df.sort_values(CONFIG.datetime_col).reset_index(drop=True)

    if df.empty:
        raise ValueError("No valid records remained after parsing dates and depth values.")

    return df


# =============================================================================
# CORE CALCULATIONS
# =============================================================================

def build_submerged_mask(df: pd.DataFrame, intervals: list[TimeInterval]) -> pd.Series:
    """Build a boolean mask identifying records inside submerged intervals."""
    mask = pd.Series(False, index=df.index)

    for interval in intervals:
        start = interval.start_ts(CONFIG.date_format)
        end = interval.end_ts(CONFIG.date_format)
        mask |= df[CONFIG.datetime_col].between(start, end, inclusive="both")

    return mask


def calculate_water_column_depth(
    h_submerged: pd.Series,
    h_mean_air: float,
    z_sensor: float,
) -> pd.Series:
    """
    Calculate corrected water-column height above the bed.

    H = h_submerged - h_mean_air + z_sensor
    """
    depth = h_submerged - h_mean_air + z_sensor
    if CONFIG.clip_negative_to_zero:
        depth = depth.clip(lower=0)
    return depth


def get_air_reference_for_period(
    df: pd.DataFrame,
    interval_index: int,
    intervals: list[TimeInterval],
    global_air_mean: float,
) -> float:
    """Select the air-reference value for one submerged period."""
    if CONFIG.air_reference_mode == "global":
        return global_air_mean

    start = intervals[interval_index].start_ts(CONFIG.date_format)
    end = intervals[interval_index].end_ts(CONFIG.date_format)

    prev_end: Optional[pd.Timestamp] = (
        None if interval_index == 0 else intervals[interval_index - 1].end_ts(CONFIG.date_format)
    )
    next_start: Optional[pd.Timestamp] = (
        None if interval_index == len(intervals) - 1 else intervals[interval_index + 1].start_ts(CONFIG.date_format)
    )

    before_mask = df[CONFIG.datetime_col] < start
    after_mask = df[CONFIG.datetime_col] > end

    if prev_end is not None:
        before_mask &= df[CONFIG.datetime_col] > prev_end
    if next_start is not None:
        after_mask &= df[CONFIG.datetime_col] < next_start

    before_dry = df.loc[before_mask, CONFIG.depth_col]
    after_dry = df.loc[after_mask, CONFIG.depth_col]

    if CONFIG.air_reference_mode == "previous_dry":
        if len(before_dry) > 0:
            return float(before_dry.mean())
        if len(after_dry) > 0:
            return float(after_dry.mean())
        return global_air_mean

    if CONFIG.air_reference_mode == "surrounding_dry":
        candidates = []
        if len(before_dry) > 0:
            candidates.append(float(before_dry.mean()))
        if len(after_dry) > 0:
            candidates.append(float(after_dry.mean()))
        return float(np.mean(candidates)) if candidates else global_air_mean

    raise ValueError(f"Unsupported air_reference_mode: {CONFIG.air_reference_mode}")


def add_moving_average(period_df: pd.DataFrame, window: str) -> pd.DataFrame:
    """Add a time-based moving average using the datetime column as index."""
    period_df = period_df.sort_values(CONFIG.datetime_col).copy()
    period_df["H_m_MA_1min"] = (
        period_df
        .set_index(CONFIG.datetime_col)["H_m"]
        .rolling(window=window, min_periods=1)
        .mean()
        .values
    )
    return period_df


def process_periods(
    df: pd.DataFrame,
    intervals: list[TimeInterval],
    z_sensor: float,
) -> tuple[pd.DataFrame, pd.DataFrame, float]:
    """Process all submerged periods and return corrected tables."""
    submerged_mask = build_submerged_mask(df, intervals)
    dry_values = df.loc[~submerged_mask, CONFIG.depth_col]

    if dry_values.empty:
        raise ValueError(
            "No dry records were found outside the defined submerged intervals. "
            "The air-reference correction requires dry records before, after, or between submerged periods."
        )

    global_air_mean = float(dry_values.mean())
    corrected_parts = []
    summary_rows = []

    for i, interval in enumerate(intervals):
        start = interval.start_ts(CONFIG.date_format)
        end = interval.end_ts(CONFIG.date_format)
        period_mask = df[CONFIG.datetime_col].between(start, end, inclusive="both")
        period_df = df.loc[period_mask, [CONFIG.datetime_col, CONFIG.depth_col]].copy()

        h_mean_air = get_air_reference_for_period(df, i, intervals, global_air_mean)

        if period_df.empty:
            summary_rows.append({
                "Period": interval.label,
                "Start": start,
                "End": end,
                "n_records": 0,
                "h_mean_air_m": h_mean_air,
                "z_sensor_m": z_sensor,
                "H_min_m": np.nan,
                "H_max_m": np.nan,
                "H_mean_m": np.nan,
                "H_MA_1min_min_m": np.nan,
                "H_MA_1min_max_m": np.nan,
                "H_MA_1min_mean_m": np.nan,
            })
            continue

        period_df["Period"] = interval.label
        period_df["h_mean_air_m"] = h_mean_air
        period_df["z_sensor_m"] = z_sensor
        period_df["H_m"] = calculate_water_column_depth(
            period_df[CONFIG.depth_col],
            h_mean_air,
            z_sensor,
        )
        period_df = add_moving_average(period_df, CONFIG.moving_average_window)
        corrected_parts.append(period_df)

        summary_rows.append({
            "Period": interval.label,
            "Start": start,
            "End": end,
            "n_records": len(period_df),
            "h_mean_air_m": h_mean_air,
            "z_sensor_m": z_sensor,
            "H_min_m": period_df["H_m"].min(),
            "H_max_m": period_df["H_m"].max(),
            "H_mean_m": period_df["H_m"].mean(),
            "H_MA_1min_min_m": period_df["H_m_MA_1min"].min(),
            "H_MA_1min_max_m": period_df["H_m_MA_1min"].max(),
            "H_MA_1min_mean_m": period_df["H_m_MA_1min"].mean(),
        })

    if not corrected_parts:
        raise ValueError("None of the configured intervals matched records in the input file.")

    corrected_df = pd.concat(corrected_parts, ignore_index=True)
    summary_df = pd.DataFrame(summary_rows)
    return corrected_df, summary_df, global_air_mean


# =============================================================================
# OUTPUTS
# =============================================================================

def save_excel(
    corrected_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    output_file: Path,
    input_file: Path,
    interval_set_label: str,
    global_air_mean: float,
) -> None:
    """Save corrected data, summaries, interval definitions, and metadata."""
    metadata_df = pd.DataFrame({
        "parameter": [
            "routine",
            "input_file",
            "interval_set_label",
            "datetime_column",
            "depth_column",
            "date_format",
            "air_reference_mode",
            "clip_negative_to_zero",
            "sensor_height_m",
            "moving_average_window",
            "global_air_mean_m",
            "plot_figsize",
            "plot_dpi",
            "plot_line_width",
            "plot_grid_alpha",
        ],
        "value": [
            "water_depth_estimation",
            input_file.name,
            interval_set_label,
            CONFIG.datetime_col,
            CONFIG.depth_col,
            CONFIG.date_format,
            CONFIG.air_reference_mode,
            CONFIG.clip_negative_to_zero,
            CONSTANTS.z_sensor,
            CONFIG.moving_average_window,
            global_air_mean,
            str(PLOT_STYLE["figsize"]),
            PLOT_STYLE["dpi"],
            PLOT_STYLE["line_width"],
            PLOT_STYLE["grid_alpha"],
        ],
    })

    intervals_df = pd.DataFrame([
        {"Period": interval.label, "Start": interval.start, "End": interval.end}
        for interval in TIME_INTERVALS
    ])

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        corrected_df.to_excel(writer, sheet_name="Corrected_H", index=False)
        summary_df.to_excel(writer, sheet_name="Period_Summary", index=False)
        metadata_df.to_excel(writer, sheet_name="Metadata", index=False)
        intervals_df.to_excel(writer, sheet_name="Defined_Intervals", index=False)


def save_single_plot(
    tmp: pd.DataFrame,
    y_col: str,
    title: str,
    output_file: Path,
    y_max: float,
) -> None:
    """Save one period plot."""
    fig, ax = plt.subplots(figsize=PLOT_STYLE["figsize"])
    ax.plot(tmp[CONFIG.datetime_col], tmp[y_col], linewidth=PLOT_STYLE["line_width"])
    ax.set_title(title)
    ax.set_xlabel(PLOT_STYLE["x_label"])
    ax.set_ylabel(PLOT_STYLE["y_label"])
    ax.set_ylim(0, y_max)
    ax.grid(True, alpha=PLOT_STYLE["grid_alpha"])
    fig.autofmt_xdate()
    output_file.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_file, dpi=PLOT_STYLE["dpi"], bbox_inches="tight")
    plt.close(fig)


def make_plots(
    corrected_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    output_dir: Path,
    y_max: float,
) -> None:
    """Generate raw and moving-average plots for each submerged period."""
    output_dir.mkdir(parents=True, exist_ok=True)

    for period in summary_df["Period"].tolist():
        tmp = corrected_df.loc[corrected_df["Period"] == period].copy()
        if tmp.empty:
            continue

        safe_period = safe_tag(period)
        raw_file = output_dir / f"{safe_period}_raw.jpg"
        ma_file = output_dir / f"{safe_period}_MA1min.jpg"

        save_single_plot(
            tmp=tmp,
            y_col="H_m",
            title=f"{period} - Corrected water depth above bed",
            output_file=raw_file,
            y_max=y_max,
        )
        save_single_plot(
            tmp=tmp,
            y_col="H_m_MA_1min",
            title=f"{period} - Corrected water depth above bed (1-minute moving average)",
            output_file=ma_file,
            y_max=y_max,
        )


# =============================================================================
# COMMAND-LINE INTERFACE
# =============================================================================

def build_parser() -> argparse.ArgumentParser:
    """Create the command-line parser."""
    parser = argparse.ArgumentParser(
        description="Calculate corrected water-column height above bed from pressure-transducer data."
    )
    parser.add_argument("--input", "-i", default=DEFAULT_INPUT_FILE, help="Input Excel or CSV file.")
    parser.add_argument("--sheet", default=None, help="Excel sheet name or index. Default: first sheet.")
    parser.add_argument("--output-excel", default=None, help="Output Excel file.")
    parser.add_argument("--output-plot-dir", default=None, help="Output directory for per-period JPG plots.")
    parser.add_argument("--sensor-height", type=float, default=CONSTANTS.z_sensor, help="Sensor height above bed, in metres.")
    parser.add_argument("--y-max", type=float, default=CONFIG.y_axis_max_m, help="Common maximum y-axis value for all plots, in metres.")
    parser.add_argument("--auto-y-max", action="store_true", help="Choose the plot y-axis maximum from the corrected data.")
    parser.add_argument("--moving-average-window", default=CONFIG.moving_average_window, help="Time-based moving-average window, for example 1min.")
    parser.add_argument(
        "--air-reference-mode",
        choices=["global", "previous_dry", "surrounding_dry"],
        default=CONFIG.air_reference_mode,
        help="How to estimate the mean air reading for each period.",
    )
    parser.add_argument(
        "--interval-set-label",
        default=DEFAULT_INTERVAL_SET_LABEL,
        help="Label identifying the defined set of submerged periods.",
    )
    parser.add_argument(
        "--keep-negative-depths",
        action="store_true",
        help="Keep negative corrected depths instead of clipping them to zero.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Execute the full processing workflow."""
    args = build_parser().parse_args(argv)

    CONFIG.air_reference_mode = args.air_reference_mode
    CONFIG.y_axis_max_m = args.y_max
    CONFIG.moving_average_window = args.moving_average_window
    CONFIG.clip_negative_to_zero = not args.keep_negative_depths
    CONSTANTS.z_sensor = args.sensor_height

    input_path = Path(args.input)
    output_excel = (
        Path(args.output_excel)
        if args.output_excel
        else output_excel_path_for(input_path, args.interval_set_label)
    )
    output_plot_dir = (
        Path(args.output_plot_dir)
        if args.output_plot_dir
        else output_plot_dir_for(input_path, args.interval_set_label)
    )

    sheet: str | int | None = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    df = read_input_data(input_path, sheet=sheet)
    corrected_df, summary_df, global_air_mean = process_periods(
        df=df,
        intervals=TIME_INTERVALS,
        z_sensor=args.sensor_height,
    )

    save_excel(
        corrected_df=corrected_df,
        summary_df=summary_df,
        output_file=output_excel,
        input_file=input_path,
        interval_set_label=args.interval_set_label,
        global_air_mean=global_air_mean,
    )
    missing_periods = summary_df.loc[summary_df["n_records"] == 0, "Period"].tolist()
    if missing_periods:
        print("Warning: no input records were found for these period(s): " + ", ".join(missing_periods))

    plot_y_max = args.y_max
    if args.auto_y_max and not corrected_df.empty:
        max_depth = float(corrected_df["H_m"].max())
        plot_y_max = max(0.1, max_depth * 1.15)

    make_plots(
        corrected_df=corrected_df,
        summary_df=summary_df,
        output_dir=output_plot_dir,
        y_max=plot_y_max,
    )

    print("=== Water-depth estimation completed ===")
    print(f"Global mean air reading = {global_air_mean:.5f} m")
    print(f"Saved corrected data to: {output_excel}")
    print(f"Saved per-period JPG plots to directory: {output_plot_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
