# Evonith Datafeed Project

## Overview
The **Evonith Datafeed Project** is a Python-based data ingestion and processing pipeline designed to:
- Download time-series data from a web interface using Selenium.
- Validate and transform the data with **Pydantic** models.
- Store the processed data in **InfluxDB**, a high-performance time-series database.
- Orchestrate the workflow using **Prefect**, ensuring seamless task execution and monitoring.

The project supports two types of variables:
1. **Temperature Variables**:
   - `BF2_BFBD Furnace Body 12975mm Temp A`
   - `BF2_BFBD Furnace Body 12975mm Temp B`
   - `BF2_BFBD Furnace Body 12975mm Temp C`
   - `BF2_BFBD Furnace Body 12975mm Temp D`
   - `BF2_BFBD Furnace Body 15162mm Temp A`
2. **Switch Variables**:
   - `BF2 Coal flow switch No 1`
   - `BF2 Coal flow switch No 10`
   - `BF2 Coal flow switch No 11`
   - `BF2 Coal flow switch No 12`
   - `BF2 Coal flow switch No 13`

## Features
- **Automated Data Download**: Retrieves data from a specified web interface at regular intervals.
- **Data Validation**: Ensures data integrity and accuracy using **Pydantic** models.
- **Time-Series Storage**: Writes validated data to InfluxDB under different measurement tags.
- **Workflow Orchestration**: Prefect handles the pipeline for downloading, processing, and writing data.

---

## Project Structure

```
project_name/
├── .env                        # Environment variables for sensitive credentials
├── src/
│   ├── config/
│   │   ├── settings.yaml        # Centralized configuration for tags and database details
│   │   ├── loader.py            # YAML loader utility
│   ├── data/
│   │   ├── downloader.py        # Selenium logic for downloading data
│   │   ├── transformer.py       # Data validation and transformation logic
│   │   ├── writer.py            # Writing data to InfluxDB
│   ├── models/
│   │   ├── data_model.py        # Pydantic models for data validation
│   ├── workflows/
│   │   ├── pipeline.py          # Prefect flow for orchestrating the pipeline
│   ├── utils/
│   │   ├── logger.py            # Logging utility for detailed logs
│   ├── app.py                   # Entry point to start the Prefect workflow
├── tests/                       # Unit tests for various modules
│   ├── test_transformer.py
├── poetry.lock                  # Poetry lock file
├── poetry.toml                  # Poetry configuration
```

---

## File Descriptions

### 1. **Environment Variables (`.env`)**
Stores sensitive credentials such as:
- Web interface username and password.
- InfluxDB authentication token.

Example:
```plaintext
USERNAME_REALTIMEDATA=your_username
PASSWORD_REALTIMEDATA=your_password
```

### 2. **Configuration Files**

#### `settings.yaml`
Contains configurable parameters such as:
- InfluxDB connection details.
- Web interface URL and download directory.
- Variable names for different measurement tags.

Example:
```yaml
influxdb:
  url: "http://localhost:8086"
  token: "your_influxdb_token"
  org: "your_org"
  bucket: "your_bucket"

selenium:
  url: "https://example.com"
  download_dir: "./downloads"

data_tags:
  temperature_variables:
    - "BF2_BFBD Furnace Body 12975mm Temp A"
    - "BF2_BFBD Furnace Body 12975mm Temp B"
    - ...
  switch_variables:
    - "BF2 Coal flow switch No 1"
    - ...
```

#### `loader.py`
Utility script to load the YAML configuration file.

---

### 3. **Data Pipeline Modules**

#### `downloader.py`
Uses Selenium to automate downloading of an Excel file from a web interface. It includes:
- WebDriver configuration.
- Automated interactions (e.g., clicking buttons, handling pop-ups).
- Returns the path of the downloaded file.

#### `transformer.py`
Processes the downloaded file:
- Reads the Excel file into a DataFrame.
- Validates and transforms data using Pydantic models.
- Groups variables by their measurement tags (e.g., `temperature_variables` and `switch_variables`).

#### `writer.py`
Writes the validated data to InfluxDB:
- Stores data under different measurement tags.
- Leverages the InfluxDB client for efficient time-series storage.

---

### 4. **Data Models (`data_model.py`)**
Defines Pydantic models for validating and transforming data:
- Ensures fields like `timestamp`, `measurement`, and `values` are properly formatted.
- Custom validation logic for handling variable-specific requirements.

---

### 5. **Prefect Workflow (`pipeline.py`)**
Defines the Prefect flow for orchestrating the pipeline:
- **Tasks**:
  - `download_task`: Downloads the Excel file.
  - `preprocess_task`: Validates and processes the data.
  - `write_task`: Writes the processed data to InfluxDB.
- **Flow**:
  - Chains tasks into a cohesive pipeline.

---

### 6. **Application Entry Point (`app.py`)**
Starts the Prefect workflow:
- Loads configuration.
- Initiates the data pipeline.

---

### 7. **Utilities**

#### `logger.py`
Sets up logging for the project. Each module uses this logger to generate detailed logs for debugging and monitoring.

---

### 8. **Unit Tests**
Located in the `tests/` directory:
- `test_transformer.py`: Tests the validation and transformation logic for handling downloaded data.

---

## Usage

### 1. **Install Dependencies**
Install dependencies using Poetry:
```bash
poetry install
```

### 2. **Run the Prefect Workflow**
Execute the data pipeline:
```bash
python src/app.py
```

### 3. **Run Unit Tests**
Run the test suite:
```bash
pytest tests/
```

---

## Future Enhancements
- Add support for additional variable types.
- Integrate real-time monitoring and alerts.
- Extend support for other time-series databases (e.g., TimescaleDB).

---

## Contributors
- **Saikumar VoptimAI**  
  [GitHub](https://github.com/saikumar-voptimai)

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.
