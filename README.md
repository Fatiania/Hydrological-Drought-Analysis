# Hydrological Drought Analysis

Event-based characterization of streamflow drought in southern Italian catchments, including perennial and intermittent flow regimes.

## Source code

### `src/daily_streamflow_screening.py`
Performs daily discharge quality control, annual completeness screening, station eligibility assessment and flow-regime classification. The analysis examines flow-duration-curve thresholds, identifies below-threshold events and calculates characteristics including duration, minimum discharge, deficit and frequency. Missing intervals and boundary-touching events are tracked separately to support interpretation of incomplete records.

### `src/threshold_pooling_comparison.py`
Evaluates threshold-dependent drought-event identification and pooling approaches. The script shares the principal quality-control and event-extraction workflow with the screening implementation but includes comparative analyses of event characteristics under alternative low-flow thresholds and independence assumptions.

### `src/hydrograph_visualization.py`
Extends the daily streamflow analysis with station-scale hydrograph and drought-diagnostic figures. Observed discharge, low-flow reference thresholds, missing intervals and identified events are displayed to examine the temporal structure of individual drought episodes.

## Data screening and thresholds

Annual records with no more than 15 missing daily values are retained; stations require at least 15 usable calendar years. Candidate low-flow thresholds are evaluated separately for perennial and intermittent stations. The reference analysis uses Q90 for perennial rivers and Q50 for intermittent rivers.

## Data

Historical daily discharge records are stored separately from the code.
