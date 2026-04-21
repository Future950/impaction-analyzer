# ImpactScan — Dental Impaction Analyzer
### Multiple Imaging Formats → Impaction Classification → SQLite Database

---

## What This Does

Reads Carestream DICOM X-ray files (panoramic, periapical), Carestream .pano files,
standard image formats (PNG, JPEG), and HTML-embedded images. Detects impacted teeth 
using radiographic heuristics, classifies them using both **Pell & Gregory** and **Winter's** 
systems, and stores everything in a structured SQLite database. A web dashboard 
(`dashboard.html`) lets you browse, filter, and export results.

---

## Supported File Formats

| Format | Extension | Source |
|--------|-----------|--------|
| DICOM | `.dcm` | Carestream/standard X-ray devices |
| Panoramic | `.pano` | Carestream panoramic images |
| PNG Image | `.png` | Any standard PNG file |
| JPEG Image | `.jpg`, `.jpeg` | Any standard JPEG file |
| HTML | `.html` | HTML with embedded images (base64 or referenced) |

---

## Files

| File | Purpose |
|---|---|
| `impaction_analyzer.py` | Main analyzer — reads DICOMs, classifies, saves DB |
| `generate_test_dicoms.py` | Generates synthetic test DICOM files |
| `dashboard.html` | Web dashboard — open in any browser |
| `dental_impactions.db` | SQLite database (created on first run) |
| `results.json` | JSON export (optional, use `--export-json`) |

---

## Installation

```bash
pip install pydicom numpy pillow openpyxl
```

---

## Usage

### 1. Process DICOM, panoramic, or image files
```bash
# Process a folder (auto-detects .dcm, .pano, .png, .jpg, .jpeg, .html)
python impaction_analyzer.py /path/to/carestream/exports/ --db dental_impactions.db

# Process specific files (mixed formats)
python impaction_analyzer.py patient1.dcm patient2.pano image.png report.html --db dental_impactions.db
```

### 2. Export results to JSON (for the dashboard)
```bash
python impaction_analyzer.py /dicoms/ --db dental_impactions.db --export-json results.json
```

### 3. Export results to Excel (for further analysis)
```bash
python impaction_analyzer.py /dicoms/ --db dental_impactions.db --export-excel results.xlsx
```
The Excel file includes:
- **Detailed Records** sheet — all impacted teeth with full classification data
- **Summary** sheet — statistics by tooth type, classification, and severity

### 4. View database summary
```bash
python impaction_analyzer.py --summary --db dental_impactions.db
```

### 5. Generate test DICOM files
```bash
python generate_test_dicoms.py
python impaction_analyzer.py test_dicoms/ --db dental_impactions.db --export-json results.json
```

---

## Open the Dashboard

Simply open `dashboard.html` in any web browser (Chrome, Firefox, Edge, Safari).

- Click **"Load Demo Dataset"** to explore immediately with synthetic data
- After running the analyzer, click **"Load results.json"** to see your real data
- The dashboard auto-loads demo data on startup

---

## Classification Systems

### Pell & Gregory (Wisdom Teeth)

**Class** — Relationship to ramus of mandible:
- **I** — Crown fully anterior to ramus; adequate space
- **II** — Crown partially covered by ramus; limited space  
- **III** — Crown fully within ramus; no eruption space

**Depth** — Relationship to occlusal plane of 2nd molar:
- **A** — Crown at or above occlusal plane
- **B** — Crown between occlusal plane and cervical line
- **C** — Crown below cervical line (deeply impacted)

### Winter's Classification (All teeth)

| Code | Name | Angle | Notes |
|---|---|---|---|
| MA | Mesioangular | +30° to +80° | Most common; risk to 2nd molar |
| DA | Distoangular | -30° to -80° | Most difficult surgically |
| V | Vertical | < ±30° | Moderate difficulty |
| H | Horizontal | > ±80° | Crown into 2nd molar root |
| T | Transverse | Bucco-lingual | CBCT required |
| IN | Inverted | 180° | Rarest; cyst association |

---

## Database Schema

```sql
patients         — Patient demographics (patient_id, name, DOB)
studies          — Each DICOM study (links to patient, file hash, date)
impacted_teeth   — Each detected impaction with full classification
raw_metadata     — DICOM tags for audit/reference
db_info          — Schema version, creation date
```

### Useful queries

```sql
-- All severely impacted wisdom teeth
SELECT p.patient_name, t.tooth_name, t.pg_class, t.pg_depth, t.winters_angle
FROM impacted_teeth t
JOIN studies s ON s.id = t.study_pk
JOIN patients p ON p.id = s.patient_pk
WHERE t.tooth_type = 'wisdom' AND t.impaction_severity = 'severe';

-- Count by classification
SELECT pg_class, pg_depth, COUNT(*) as n
FROM impacted_teeth
WHERE tooth_type = 'wisdom'
GROUP BY pg_class, pg_depth
ORDER BY pg_class, pg_depth;

-- Patients with multiple impactions
SELECT p.patient_id, p.patient_name, COUNT(*) as impactions
FROM impacted_teeth t
JOIN studies s ON s.id = t.study_pk
JOIN patients p ON p.id = s.patient_pk
GROUP BY p.id
HAVING impactions > 2
ORDER BY impactions DESC;
```

---

## Important Notes

**Heuristic Detection:** This analyzer uses validated radiographic heuristics
(density gradients, gradient energy, apical/coronal density ratios) to detect
impaction. It is not a replacement for clinical radiographic interpretation.

**For Production Use:** Integrate a CNN model trained on labelled panoramic
radiographs for higher sensitivity and specificity. The architecture is designed
to accept a trained model as a drop-in replacement for the heuristic classifier.

**DICOM Compatibility:** Tested with standard Carestream DX modality exports.
The analyzer handles malformed DICOM headers gracefully using `force=True`.

**Duplicate Prevention:** Each file is SHA-256 hashed on import. Re-running
the analyzer on the same files will not create duplicate records.

---

## Tooth Type Categories

| Type | FDI Numbers | Classification |
|---|---|---|
| Wisdom | 18, 28, 38, 48 | Pell & Gregory + Winter's |
| Canine | 13, 23, 33, 43 | Winter's angulation |
| Premolar | 14, 15, 24, 25, 34, 35, 44, 45 | Winter's angulation |
| Other | All remaining | Winter's angulation |

FDI numbering system used throughout (ISO 3950).
