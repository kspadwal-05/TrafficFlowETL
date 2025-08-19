# TrafficFlow ETL

**Tech used:** Docker, MS Access, Pandas, Python, SQL, VBA

This project takes traffic device data from **Microsoft Access**, cleans it with **Pandas**, saves it into a **SQL database**, and sends the results as JSON to a **REST API**.  
If you already export JSON from Access using the provided **VBA script**, the pipeline will use that as the main source.

## How it works

1. **Extract (Get the data)**  
   - First choice: use `data/input/export_from_vba.json` (created by the VBA JSON Generator).  
   - If that file isn’t there, try to open the Access database (`traffic_database.accdb`) directly.  
   - If Access isn’t available, it falls back to a small `sample.csv` so the pipeline can still run.

2. **Transform (Clean the data)**  
   - Uses Pandas to make column names and formats consistent.  
   - Fixes street names and intersections using fuzzy matching (compareStreets logic).  
   - If data already comes from the VBA JSON, the column names are kept as is, just adjusted slightly for loading.

3. **Load (Save the data)**  
   - Saves everything into a SQLite database at `data/processed/trafficflow.db`.  
   - Table name: `traffic_devices`.  
   - Duplicate records are updated automatically.  
   - Rows missing an ID are skipped.

4. **Publish (Send the data)**  
   - The cleaned data is written to `data/processed/transformed_data.json`.  
   - That JSON is sent to the API endpoint set in your `.env` file.
---