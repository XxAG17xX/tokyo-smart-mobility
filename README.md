# Tokyo smart mobility: evidence for our case study

CS7NS4 Urban Computing, Trinity College Dublin, Assignment 1 (Tokyo). This repository holds the data, code and figures behind the **Mobility and Transport** section of our report, so every number in it can be checked.

All transport data comes from [ODPT](https://developer.odpt.org/) (Toei Bureau of Transportation, CC BY 4.0).

## Claim → evidence

| Claim in the report | Evidence | How it was made |
|---|---|---|
| Toei's live feeds need no key and update every 10–30 s; ~100 trains and ~600 buses in the morning rush | [`api-samples/`](api-samples) (raw `.pb` + readable `.json`, 3 runs) | `fetch_toei.py` |
| The feeds contain **no delay field** | any `toei_train_trip_update_*.json`, e.g. [30 Sep 07:22](api-samples/toei_train_trip_update_20260930-072259.json) | search the file for `"delay"`: no matches |
| Train alerts feed was empty in all 3 runs | [`toei_train_alert_*.json`](api-samples) | `fetch_toei.py` |
| At 07:22 on 30 Sep 2026, at least 192 of 650 buses (30%) were 3+ min late | [`bus_delays_20260930-072256.csv`](api-samples/bus_delays_20260930-072256.csv) (one row per bus) | `bus_delays.py api-samples/toei_bus_vehicle_20260930-072259.pb` |
| Late buses cluster in the east (Kinshichō, Monzen-nakachō) | [static map](figures/fig_bus_map.png) · **[interactive map in kepler.gl](KEPLER_LINK)** | `bus_map.py`; [kepler map file](kepler/toei_bus_delays_20260930_0722.json) |
| ODPT as a city platform; the proposed loop | [`diagrams/`](diagrams) (draw.io files, open and edit in diagrams.net) | drawn in [diagrams.net](https://app.diagrams.net/) |

## Reproduce

```bash
pip install -r requirements.txt
python fetch_toei.py                                            # live snapshot of the 4 Toei feeds
python bus_delays.py api-samples/toei_bus_vehicle_20260930-072259.pb   # delays for a saved snapshot
python bus_map.py api-samples/bus_delays_20260930-072256.csv     # map
```

`bus_delays.py` compares each bus's live position with Toei's published timetable ([`ToeiBus-GTFS_20261006.zip`](api-samples/ToeiBus-GTFS_20261006.zip), version of 6 Oct 2026). Delays are **lower bounds**: each bus is compared with the stop it is heading to.

## AI use

The scripts, diagrams and map files were written with the help of Claude (Anthropic) and checked by us. See the AI declaration in the report.

## Credits

Data: ODPT / Tokyo Metropolitan Bureau of Transportation, CC BY 4.0. Base map in `bus_map.py`: Esri, HERE, Garmin, © OpenStreetMap contributors.
