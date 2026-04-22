import sqlite3
import json
from pathlib import Path

DB_FILE = "vehicle_oils.db"


# -----------------------------
# CREATE SINGLE TABLE
# -----------------------------
def create_tables():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("DROP TABLE IF EXISTS vehicle_oils")

    cursor.execute("""
    CREATE TABLE vehicle_oils (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT,
        year INTEGER,
        make TEXT,
        model TEXT,
        engine TEXT,
        oil_type TEXT,
        recommendation_level TEXT,
        temperature TEXT,
        capacity_with_filter_quarts REAL,
        capacity_with_filter_liters REAL,
        capacity_without_filter_quarts REAL,
        capacity_without_filter_liters REAL
    )
    """)

    conn.commit()
    conn.close()
    print("✓ Table created successfully")


# -----------------------------
# INSERT FLATTENED DATA
# -----------------------------
def insert_flat_data(cursor, filename, vehicle_info, engines):
    year = vehicle_info.get("year")
    make = vehicle_info.get("make")
    model = vehicle_info.get("model")

    for engine_name, engine_data in engines.items():

        capacity = engine_data.get("oil_capacity", {})
        with_filter = capacity.get("with_filter", {}) or {}
        without_filter = capacity.get("without_filter", {}) or {}

        oil_recs = engine_data.get("oil_recommendations", [])

        for rec in oil_recs:
            oil_type = rec.get("oil_type")
            level = rec.get("recommendation_level")
            temps = rec.get("temperature_condition", [])

            # If no temperature, still insert one row
            if not temps:
                temps = ["N/A"]

            for temp in temps:
                cursor.execute("""
                INSERT INTO vehicle_oils (
                    source_file, year, make, model, engine,
                    oil_type, recommendation_level, temperature,
                    capacity_with_filter_quarts, capacity_with_filter_liters,
                    capacity_without_filter_quarts, capacity_without_filter_liters
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    filename,
                    year,
                    make,
                    model,
                    engine_name,
                    oil_type,
                    level,
                    temp,
                    with_filter.get("quarts"),
                    with_filter.get("liters"),
                    without_filter.get("quarts"),
                    without_filter.get("liters")
                ))


# -----------------------------
# MAIN MIGRATION FUNCTION
# -----------------------------
def migrate_json_to_sqlite(json_file):

    # Load JSON
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print("✗ JSON file not found")
        return
    except json.JSONDecodeError:
        print("✗ Invalid JSON format")
        return

    # Create table
    create_tables()

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    try:
        for filename, vehicle_data in data.items():

            print(f"Processing: {filename}")

            vehicle_info = vehicle_data.get("Vehicle", {})
            engines = vehicle_data.get("engines", {})

            insert_flat_data(cursor, filename, vehicle_info, engines)

        conn.commit()
        print("\n✓ Data inserted successfully!")
        print(f"✓ Database saved as: {DB_FILE}")

    except Exception as e:
        conn.rollback()
        print("✗ Error:", e)

    finally:
        cursor.close()
        conn.close()


# -----------------------------
# RUN SCRIPT
# -----------------------------
if __name__ == "__main__":

    print("=" * 50)
    print("Vehicle Oil Database (Single Table)")
    print("=" * 50)

    # Auto-detect JSON file
    possible_paths = [
        "structured_results.json",
        Path(__file__).parent / "structured_results.json"
    ]

    json_file = None
    for path in possible_paths:
        if Path(path).exists():
            json_file = str(path)
            break

    if not json_file:
        print("✗ structured_results.json not found")
    else:
        print(f"✓ Using: {json_file}")
        migrate_json_to_sqlite(json_file)