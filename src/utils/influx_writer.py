from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS
import pandas as pd
from collections import defaultdict

def push_dataframe_to_influx(df, bucket, measurement, influx_config, field_mapping=None):
    client = InfluxDBClient(
        url=influx_config["url"],
        token=influx_config["token"],
        org=influx_config["org"]
    )
    write_api = client.write_api(write_options=SYNCHRONOUS)

    timestamp_col = influx_config.get("timestamp_col", "date")
    if timestamp_col not in df.columns:
        print(f"❌ Timestamp column '{timestamp_col}' not found.")
        return

    # Ensure mapped fields exist
    if field_mapping:
        for label in field_mapping.values():
            if label != timestamp_col and label not in df.columns:
                df[label] = pd.NA

    allowed_fields = set(field_mapping.values()) if field_mapping else set(df.columns)
    allowed_fields.discard(timestamp_col)
    df = df[[c for c in df.columns if c in allowed_fields or c == timestamp_col]]

    # Track status
    write_success = defaultdict(int)
    write_skipped = defaultdict(list)

    for _, row in df.iterrows():
        timestamp = row.get(timestamp_col)
        if pd.isna(timestamp):
            continue

        point = Point(measurement).time(pd.to_datetime(timestamp), WritePrecision.NS)

        for col in allowed_fields:
            val = row.get(col, pd.NA)
            if pd.isna(val):
                write_skipped[col].append("NaN")
                continue
            try:
                point = point.field(col, float(val))
                write_success[col] += 1
            except Exception:
                write_skipped[col].append("ValueError")

        write_api.write(bucket=bucket, record=point)

    client.close()

    # Summary
    print("\n📌 InfluxDB Write Summary:")
    print("✅ Written at least once:")
    for col in sorted(write_success.keys()):
        print(f"  - {col}: {write_success[col]} rows")

    skipped = sorted(set(allowed_fields) - set(write_success.keys()))
    if skipped:
        print("\n⚠️ Skipped columns with reasons:")
        for col in skipped:
            reasons = write_skipped.get(col, [])
            reason_counts = {r: reasons.count(r) for r in set(reasons)}
            print(f"  - {col}: {reason_counts if reason_counts else 'No values in any row'}")
