-----

# Vehicle Manual Oil Spec Extractor

A robust Python script designed to automatically parse vehicle owner's manuals (PDFs) for engine oil specifications. Using advanced regular expressions, proximity heuristics, and conditional OCR (Optical Character Recognition), it extracts vehicle details, engine sizes, engine types, oil capacities (with and without filters), recommended oil viscosities, and specific temperature conditions.

## Features

  * **Smart PDF Parsing:** Evaluates whether a PDF has native text or is scanned. Automatically routes image-heavy pages through Tesseract OCR.
  * **Context-Aware Extraction:** Uses surrounding text context to differentiate engine oil capacities from other vehicle fluids (e.g., transmission fluid, coolant).
  * **Engine & Oil Mapping:** Intelligently pairs engine displacements (e.g., "3.5L V6") with their specific oil capacities and recommended viscosities.
  * **Temperature Condition Parsing:** Extracts and normalizes temperature ranges associated with conditional oil viscosities (e.g., "5W-20 for cold weather").
  * **Structured JSON Output:** Exports all findings into a clean, hierarchical JSON file for easy integration into databases or applications.

-----

## Prerequisites

### 1\. System Requirements

You must have **Tesseract OCR** installed on your system for the image-to-text processing to work on scanned manuals.

  * **macOS:** `brew install tesseract`
  * **Ubuntu/Debian:** `sudo apt-get install tesseract-ocr`
  * **Windows:** Download the installer from the [UB-Mannheim Tesseract GitHub](https://www.google.com/search?q=https://github.com/UB-Mannheim/tesseract/wiki).

### 2\. Python Dependencies

Install the required Python packages using pip:

```bash
pip install PyMuPDF Pillow pytesseract
```

-----

## Configuration & Setup

Open the script and update the following variables at the top of the file if needed:

  * `OUTPUT_FILE`: Desired name for the final output file (default: `structured_results.json`).
  * `REFERENCE_FILE`: (Optional) A local JSON file (`vehicle_reference.json`) formatted as `{"Make": ["Model1", "Model2"]}` used to help normalize text when detecting vehicle make/models.

-----

## Usage

Ensure your target PDFs are located where the script expects them, then run the script from your terminal:

```bash
python reader.py
```

The script will output its progress to the console, showing which files are being processed, whether it is using direct text extraction or OCR, and what engines it has found.

-----

## Output Format

The script generates a `structured_results.json` file. The output is mapped by the original PDF filename.

**Example Output:**

```json
{
    "2018-Honda-Civic.pdf": {
        "Vehicle": {
            "year": 2018,
            "make": "Honda",
            "model": "Civic",
            "engine_types": ["I4", "TURBO"],
            "displayName": "2018 Honda Civic"
        },
        "engines": {
            "1.5L Turbo": {
                "oil_capacity": {
                    "with_filter": {
                        "quarts": 3.7,
                        "liters": 3.5
                    },
                    "without_filter": null
                },
                "oil_recommendations": [
                    {
                        "oil_type": "0W-20",
                        "recommendation_level": "primary",
                        "temperature_condition": ["all temperatures"]
                    }
                ]
            }
        }
    }
}
```

-----

## Deep Dive: How the Code Works

The script operates through a multi-stage pipeline designed to handle the highly inconsistent formatting found in automotive manuals. Here is a detailed breakdown of the logic:

### Phase 1: Document Triage & Text Extraction

1.  **Character Density Check:** The script loads the PDF via `PyMuPDF` (`fitz`) and calculates the average number of characters per page.
2.  **Routing:** If the average is above 800 characters, it assumes the PDF contains native, selectable text and uses fast direct extraction. If the average is below 800, it assumes the manual is composed of scanned images and routes the document to `pytesseract` to OCR the pages containing images.

### Phase 2: Metadata Identification

1.  **Filename Parsing:** It first attempts to extract the Year, Make, and Model via Regex directly from the filename (e.g., parsing `2017-Ford-F150.pdf`).
2.  **Prose Fallback:** If the filename is generic, it reads the first 5 pages of the document. It tokenizes words, strips out common manual stop-words ("guide", "manual", "page"), looks for 4-digit years, and searches for common body styles (sedan, truck, suv) to infer the correct Make and Model.

### Phase 3: Engine Extraction (The Two-Pass System)

Extracting engines is notoriously difficult because fluid capacities (e.g., "5.7L of coolant") look exactly like engine sizes (e.g., "5.7L V8"). The script handles this safely:

1.  **Pass 1 (Structured Tables):** It scans for table headers containing words like "Engine", "VIN", "Spark Gap", or "Code". When it finds a table, it extracts engine sizes (e.g., `1.5L`) and looks at the adjacent text in the row to find mechanical variants (e.g., `Turbo`, `V6`, `I4`). It explicitly ignores rows containing keywords like `coolant`, `payload`, or `differential`.
2.  **Pass 2 (Prose Scanning):** If no tables are found, it splits the document into sentences. It looks for numbers matching engine patterns (0.6L to 12.0L) that are surrounded by engine-specific context words ("displacement", "horsepower", "cylinder").
3.  **Outlier Filtering:** To prevent false positives, it clusters detected engines by size. If it finds a bizarre outlier (e.g., a random "7.5L" mention in a manual where all other engines are around "2.0L"), it trims the outlier assuming it was a misread part number or capacity.

### Phase 4: Fluid Capacity Mapping

1.  **Locating the Oil Section:** The script searches line-by-line for headers explicitly stating "Engine Oil", "Oil Capacity", or "Oil with Filter".
2.  **Extraction:** Once inside the oil section, it searches for capacity strings (`quarts`, `qts`, `liters`) using Regex. It converts everything to standard floats to ensure numerical limits make sense (rejecting values outside 1.0 to 20.0 quarts).
3.  **Reconciliation:** It attempts to match a detected capacity on a line with an engine size mentioned on that exact same line. If the manual only lists a single, shared oil capacity for all engines, it falls back to a broad document scan and applies that "shared capacity" to all previously detected engines.

### Phase 5: Oil Viscosity & Temperature Analysis (The Scoring Algorithm)

Manuals often list many oil weights (e.g., "0W-20 is best", "5W-30 may be used in warm climates"). The script uses a weighted point system to figure out the primary recommendation:

1.  **Blacklisting (Pass 0):** It searches for phrases like "do not use" and permanently blacklists any oil weight mentioned immediately after.
2.  **Inline Links (Pass 1):** If an oil is explicitly paired with an engine in parentheses `1.5L (SAE 0W-20)`, it is heavily weighted and tied strictly to that specific engine.
3.  **Primary Signals (+10 points):** Regex searches for strong definitive phrases like "is the best viscosity" or "preferred engine oil" next to an oil weight.
4.  **Secondary Signals (+2 to +6 points):** Regex catches alternative phrases like "may be used", "is acceptable", or "recommended".
5.  **Temperature Parsing:** Surrounding sentences are fed into `extract_temperature()`. The script converts Celsius to Fahrenheit, normalizes ranges, and categorizes them. For example, if it detects "below 40°F", it attaches a "cold weather" tag strictly to the oil mentioned in that sentence.

### Phase 6: Synthesis & Output

Finally, the script builds the JSON tree. It maps the highest-scoring oil as `primary` and lower-scoring oils as `secondary`. It pairs the final capacities (both with and without filter) to the correct engine variants, ensures consistent naming, and dumps the result into `structured_results.json`.