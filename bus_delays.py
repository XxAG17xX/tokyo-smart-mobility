"""Estimate how late Toei buses are, using two free ODPT files (no key needed).

The live feed only says WHERE each bus is. The timetable says WHEN it should be there.
Comparing the two gives the delay that the live feed does not report.

Run:  python bus_delays.py                         (live data)
      python bus_delays.py api-samples/<file>.pb    (a bus feed saved earlier by fetch_toei.py)
"""
import csv, io, statistics, sys, urllib.request, zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from google.transit import gtfs_realtime_pb2

LIVE = "https://api-public.odpt.org/api/v4/gtfs/realtime/ToeiBus"
TIMETABLE = "https://api-public.odpt.org/api/v4/files/Toei/data/ToeiBus-GTFS.zip"
TOKYO = timezone(timedelta(hours=9))

# 1. Where is every bus right now? (live GTFS-Realtime feed)
feed = gtfs_realtime_pb2.FeedMessage()
feed.ParseFromString(open(sys.argv[1], "rb").read() if len(sys.argv) > 1 else urllib.request.urlopen(LIVE).read())
buses = [e.vehicle for e in feed.entity if e.HasField("vehicle")]
running = {b.trip.trip_id for b in buses}

# 2. When should each bus be at each stop? (static GTFS timetable, a zip of CSV files)
timetable = zipfile.ZipFile(io.BytesIO(urllib.request.urlopen(TIMETABLE).read()))
scheduled = {}  # (trip, stop number) -> "HH:MM:SS"
with timetable.open("stop_times.txt") as f:
    for row in csv.DictReader(io.TextIOWrapper(f, "utf-8-sig")):
        if row["trip_id"] in running:  # only keep trips that are running now
            scheduled[row["trip_id"], int(row["stop_sequence"])] = row["arrival_time"]

# 3. Delay = time the bus was seen heading to its next stop - time it was due there.
#    The bus has not reached that stop yet, so the real delay is at least this big.
rows = []
for b in buses:
    due = scheduled.get((b.trip.trip_id, b.current_stop_sequence))
    if not due:
        continue  # trip not found in the timetable
    seen = datetime.fromtimestamp(b.timestamp, TOKYO)
    h, m, s = map(int, due.split(":"))
    due_time = seen.replace(hour=0, minute=0, second=0) + timedelta(hours=h, minutes=m, seconds=s)
    if due_time - seen > timedelta(hours=12):  # trip started yesterday (after midnight)
        due_time -= timedelta(days=1)
    delay_min = (seen - due_time).total_seconds() / 60
    rows.append({"bus": b.vehicle.id, "route": b.trip.route_id, "seen": seen.strftime("%H:%M"),
                 "delay_min": round(delay_min, 1), "lat": b.position.latitude, "lon": b.position.longitude})

# 4. Summary + save
delays = [r["delay_min"] for r in rows]
stamp = datetime.fromtimestamp(feed.header.timestamp, TOKYO).strftime("%Y%m%d-%H%M%S")
out = Path(__file__).parent / "api-samples" / f"bus_delays_{stamp}.csv"
with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["bus", "route", "seen", "delay_min", "lat", "lon"])
    w.writeheader()
    w.writerows(rows)

print(f"Tokyo time {stamp} | buses in live feed: {len(buses)} | matched to timetable: {len(rows)}")
if delays:
    print(f"median delay: {statistics.median(delays):.1f} min")
    for limit in (1, 3, 5, 10):
        print(f"  at least {limit:>2} min late: {sum(d >= limit for d in delays)} buses")
print(f"saved {out.name}")
