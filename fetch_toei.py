"""Download Toei's live GTFS-Realtime feeds from ODPT, save them, and print a summary.

Setup once:  python -m pip install gtfs-realtime-bindings
Run:         python fetch_toei.py
"""
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from google.protobuf.json_format import MessageToJson
from google.transit import gtfs_realtime_pb2

# Public ODPT links, no key needed (CC BY 4.0)
FEEDS = {
    "toei_train_vehicle": "https://api-public.odpt.org/api/v4/gtfs/realtime/toei_odpt_train_vehicle",
    "toei_train_trip_update": "https://api-public.odpt.org/api/v4/gtfs/realtime/toei_odpt_train_trip_update",
    "toei_train_alert": "https://api-public.odpt.org/api/v4/gtfs/realtime/toei_odpt_train_alert",
    "toei_bus_vehicle": "https://api-public.odpt.org/api/v4/gtfs/realtime/ToeiBus",
}
TOKYO = timezone(timedelta(hours=9))  # Tokyo is UTC+9

folder = Path(__file__).parent / "api-samples"
folder.mkdir(exist_ok=True)
stamp = datetime.now(TOKYO).strftime("%Y%m%d-%H%M%S")

for name, url in FEEDS.items():
    # 1. Download. The data arrives in a compact binary format (protobuf).
    raw = urllib.request.urlopen(url).read()

    # 2. Decode it with Google's official GTFS-Realtime library.
    feed = gtfs_realtime_pb2.FeedMessage()
    feed.ParseFromString(raw)

    # 3. Save the original file and a readable JSON copy.
    text = MessageToJson(feed, ensure_ascii=False)
    (folder / f"{name}_{stamp}.pb").write_bytes(raw)
    (folder / f"{name}_{stamp}.json").write_text(text, encoding="utf-8")

    # 4. Summary: how many trains/buses/updates, is there any delay data, when was it made.
    print(name)
    print("  items:", len(feed.entity))
    print("  has a delay field:", '"delay":' in text)
    print("  feed time (Tokyo):", datetime.fromtimestamp(feed.header.timestamp, TOKYO))
