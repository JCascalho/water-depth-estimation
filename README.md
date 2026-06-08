# Water Depth Estimation

Python routine for estimating corrected water-column height above the seabed from pressure-transducer records.

The routine reads pressure-transducer time series, selects submerged intervals, estimates an air-reference offset, calculates corrected water depth, applies a time-based moving average, and exports tables and figures for reproducible analysis.

## Features

- Reads Excel or CSV pressure-transducer input files.
- Uses predefined submerged intervals configured inside the script.
- Supports global, previous-dry, or surrounding-dry air-reference correction.
- Calculates corrected water-column height above the bed.
- Applies a configurable time-based moving average.
- Exports a structured Excel workbook with corrected data, summary statistics, interval definitions, and metadata.
- Generates raw and smoothed JPG plots for each submerged interval.

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Input Format

The input file can be CSV, XLSX, or XLS. It must contain at least:

- `Date_Time`: date and time of each record
- `Depth_PT_m`: pressure-transducer depth reading in metres

Example:

```csv
Date_Time,Depth_PT_m
24-09-2022 11:50:00,0.012
24-09-2022 12:13:00,0.326
24-09-2022 12:14:00,0.331
24-09-2022 17:00:58,0.298
24-09-2022 17:20:00,0.014
```

## Submerged Intervals

The submerged periods are defined in the `TIME_INTERVALS` block inside
`water_depth_estimation.py`.

Example:

```python
TIME_INTERVALS = [
    TimeInterval("24-09-2022 12:13:00", "24-09-2022 17:00:58", "1stPeriod"),
    TimeInterval("25-09-2022 00:50:00", "25-09-2022 04:54:58", "2ndPeriod"),
]
```

For a new campaign, edit only this block and keep the rest of the routine unchanged.

## Usage

CSV input:

```bash
python water_depth_estimation.py --input examples/example_pressure_transducer_data.csv --output-excel results/water_depth_results.xlsx --output-plot-dir results/plots --interval-set-label example --auto-y-max
```

Excel input:

```bash
python water_depth_estimation.py --input INPUT_PT_DATA.xlsx --sheet Sheet1 --interval-set-label September_2022
```

Spyder example:

```python
%runfile "C:/path/to/water_depth_estimation.py" --wdir "C:/path/to/water-depth-estimation" --args "--input examples/example_pressure_transducer_data.csv --output-excel results/water_depth_results.xlsx --output-plot-dir results/plots --interval-set-label example --auto-y-max"
```

## Outputs

The output Excel workbook includes:

- `Corrected_H`: corrected water-depth time series
- `Period_Summary`: per-period summary statistics
- `Defined_Intervals`: interval definitions used in the run
- `Metadata`: processing settings and input metadata

The plot directory includes:

- one raw corrected-depth plot per interval
- one moving-average plot per interval

## Notes

The corrected water depth is calculated as:

```text
H = pressure_depth_m - air_reference_m + sensor_height_m
```

By default, negative corrected depths are clipped to zero. Use `--keep-negative-depths` if you want to preserve negative values.

## Citation

If you use this routine, please cite the archived GitHub release DOI generated through Zenodo.
