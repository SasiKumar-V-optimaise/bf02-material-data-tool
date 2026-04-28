import pandas as pd
from influxdb_client import InfluxDBClient, WriteOptions

# ---------------- USER SETTINGS ----------------
FILE = r"C:\Users\sasik\Downloads\Book1.xlsx"
BUCKET = "bf2_evonith_offline_utc"
MEASUREMENT = "hotmetal_slag_data"
ORG = "Blast Furnace, Evonith"
TOKEN = "yT-IxzwVJdbagjlJM0yByybQA83IvWHkew4cy97TcMs1BSFSm8bAPMOyoIrKir06M7xo3s5xV6YEHe7jdFnBLw=="
URL = "https://eu-central-1-1.aws.cloud2.influxdata.com"
# -----------------------------------------------

# 1️⃣ Read Excel file
df = pd.read_excel(FILE)

# 2️⃣ Detect timestamp column
time_col = [c for c in df.columns if "time" in c.lower()][0]

# 3️⃣ Build dataframe for InfluxDB
df_influx = pd.DataFrame()
df_influx["_measurement"] = MEASUREMENT
df_influx["time"] = pd.to_datetime(df[time_col])

# Add remaining columns as fields
for col in df.columns:
    if col != time_col:
        df_influx[col] = df[col]

print("Preview of data prepared for upload:")
print(df_influx.head())

# 4️ Write to InfluxDB
with InfluxDBClient(url=URL, token=TOKEN, org=ORG) as client:
    write_api = client.write_api(write_options=WriteOptions(batch_size=2500))
    write_api.write(
        bucket=BUCKET,
        record=df_influx,
        data_frame_measurement_name="_measurement",
        data_frame_timestamp_column="time"
    )

print("\n DONE — Data uploaded to InfluxDB successfully!")
