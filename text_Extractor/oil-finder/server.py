import os
import json
from flask import Flask, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# PATH TO STRUCTURED RESULTS JSON
JSON_FILE = os.path.join(os.path.dirname(__file__), "..", "textExtraction", "structured_results.json")

print(f"JSON File path: {JSON_FILE}")
print(f"File exists: {os.path.exists(JSON_FILE)}")


def load_vehicles_from_json():
    """Load vehicles data from structured_results.json"""
    try:
        if not os.path.exists(JSON_FILE):
            print(f"ERROR: JSON file not found at {JSON_FILE}")
            return []
        
        with open(JSON_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"Loaded JSON with {len(data)} documents")
        
        vehicles_list = []
        
        for doc_name, doc_data in data.items():
            vehicle_info = doc_data.get("Vehicle", {})
            engines = doc_data.get("engines", {})
            
            print(f"Processing {doc_name}: {len(engines)} engines")
            
            for engine_size, engine_data in engines.items():
                vehicle = {
                    "year": vehicle_info.get("year"),
                    "make": vehicle_info.get("make"),
                    "model": vehicle_info.get("model"),
                    "engine": engine_size,
                    "displayName": vehicle_info.get("displayName"),
                    "capacity": engine_data.get("oil_capacity", {}),
                    "oils": []
                }
                
                # Add oil recommendations
                for rec in engine_data.get("oil_recommendations", []):
                    vehicle["oils"].append({
                        "oil_type": rec.get("oil_type"),
                        "recommendation_level": rec.get("recommendation_level"),
                        "temperature": rec.get("temperature_condition", [])
                    })
                
                vehicles_list.append(vehicle)
        
        print(f"Total vehicles loaded: {len(vehicles_list)}")
        return vehicles_list
    except Exception as e:
        print(f"Error loading JSON: {e}")
        import traceback
        traceback.print_exc()
        return []


@app.route("/", methods=["GET"])
def index():
    """Serve the frontend"""
    return jsonify({"message": "Oil Finder API is running"}), 200


@app.route("/api/vehicles", methods=["GET"])
def get_vehicles():
    vehicles = load_vehicles_from_json()
    return jsonify(vehicles)


if __name__ == "__main__":
    app.run(debug=True)