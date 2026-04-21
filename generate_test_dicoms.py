"""
generate_test_dicoms.py
Creates synthetic DICOM panoramic-like images for testing the impaction analyzer.
Run this once to produce sample .dcm files.
"""

import os
import numpy as np
import pydicom
from pydicom.dataset import Dataset, FileMetaDataset
from pydicom.uid import generate_uid, ExplicitVRLittleEndian
import datetime

OUTPUT_DIR = "test_dicoms"
os.makedirs(OUTPUT_DIR, exist_ok=True)

PATIENTS = [
    {"id": "PT001", "name": "Asante^Kofi",   "dob": "19900315", "scenario": "wisdom_impacted"},
    {"id": "PT002", "name": "Mensah^Abena",  "dob": "20011122", "scenario": "canine_impacted"},
    {"id": "PT003", "name": "Boateng^Kwame", "dob": "19851005", "scenario": "multiple"},
    {"id": "PT004", "name": "Adjei^Ama",     "dob": "20050630", "scenario": "normal"},
    {"id": "PT005", "name": "Owusu^Yaw",     "dob": "19950818", "scenario": "severe_multiple"},
]


def make_panoramic(scenario: str) -> np.ndarray:
    """Return a 512x1024 float32 array mimicking a panoramic x-ray."""
    img = np.random.normal(80, 15, (512, 1024)).astype(np.float32)
    img = np.clip(img, 0, 255)

    # Add tooth-like bright ellipses across the arch
    def tooth(cx, cy, rx, ry, intensity=200, angle_deg=0):
        for y in range(max(0, cy-ry-5), min(512, cy+ry+5)):
            for x in range(max(0, cx-rx-5), min(1024, cx+rx+5)):
                # Rotate
                rad = np.radians(angle_deg)
                dx, dy = x - cx, y - cy
                rx2 = dx*np.cos(rad) + dy*np.sin(rad)
                ry2 = -dx*np.sin(rad) + dy*np.cos(rad)
                if (rx2/max(rx,1))**2 + (ry2/max(ry,1))**2 <= 1:
                    img[y, x] = min(255, img[y, x] + intensity)

    # Normal erupted teeth — upper arch
    upper_cx = [80, 160, 240, 320, 400, 480, 560, 640, 720, 800, 880, 960,
                100, 200, 300, 500, 600, 850]
    for i, cx in enumerate(upper_cx):
        tooth(cx, 160, 28, 45, angle_deg=0)

    # Normal erupted teeth — lower arch
    lower_cx = [80, 160, 240, 320, 400, 480, 560, 640, 720, 800, 880, 960]
    for cx in lower_cx:
        tooth(cx, 360, 28, 45, angle_deg=0)

    if scenario == "wisdom_impacted":
        # Upper right wisdom (FDI 18) — mesioangular, high density, partially covered
        tooth(60, 130, 35, 25, intensity=220, angle_deg=45)   # mesioangular UR
        # Lower left wisdom (FDI 38) — horizontal
        tooth(750, 390, 40, 20, intensity=230, angle_deg=90)

    elif scenario == "canine_impacted":
        # Upper left canine (FDI 23) — high in arch, vertical but unerupted
        tooth(520, 110, 22, 55, intensity=240, angle_deg=5)
        # Lower right canine (FDI 43) — transverse
        tooth(870, 370, 30, 18, intensity=215, angle_deg=70)

    elif scenario == "multiple":
        # Wisdom teeth — both sides impacted
        tooth(55,  125, 33, 22, intensity=225, angle_deg=50)   # 18 mesioangular
        tooth(975, 125, 33, 22, intensity=225, angle_deg=-40)  # 28 distoangular
        tooth(60,  395, 38, 18, intensity=230, angle_deg=88)   # 38 horizontal
        # Canine
        tooth(310, 105, 20, 52, intensity=240, angle_deg=8)    # 13 vertical

    elif scenario == "severe_multiple":
        # All four third molars impacted + bilateral canines
        for cx, cy, ang in [
            (55,  120, 55), (975, 120, -55),   # upper wisdom
            (60,  400, 90), (975, 400, -90),   # lower wisdom
            (310, 100, 10), (720, 100, -10),   # upper canines
        ]:
            tooth(cx, cy, 35, 22, intensity=235, angle_deg=ang)

    # scenario == "normal" → no additional impacted teeth

    return np.clip(img, 0, 255).astype(np.uint16)


def write_dicom(patient: dict, idx: int):
    meta = FileMetaDataset()
    meta.MediaStorageSOPClassUID    = "1.2.840.10008.5.1.4.1.1.1"
    meta.MediaStorageSOPInstanceUID = generate_uid()
    meta.TransferSyntaxUID          = ExplicitVRLittleEndian
    meta.is_implicit_VR             = False
    meta.is_little_endian           = True

    ds = Dataset()
    ds.file_meta            = meta
    ds.is_implicit_VR       = False
    ds.is_little_endian     = True

    ds.PatientID            = patient["id"]
    ds.PatientName          = patient["name"]
    ds.PatientBirthDate     = patient["dob"]
    ds.StudyDate            = datetime.date.today().strftime("%Y%m%d")
    ds.StudyTime            = "120000"
    ds.Modality             = "DX"
    ds.StudyInstanceUID     = generate_uid()
    ds.SeriesInstanceUID    = generate_uid()
    ds.SOPInstanceUID       = meta.MediaStorageSOPInstanceUID
    ds.SOPClassUID          = meta.MediaStorageSOPClassUID
    ds.Rows                 = 512
    ds.Columns              = 1024
    ds.BitsAllocated        = 16
    ds.BitsStored           = 12
    ds.HighBit              = 11
    ds.PixelRepresentation  = 0
    ds.SamplesPerPixel      = 1
    ds.PhotometricInterpretation = "MONOCHROME2"
    ds.InstitutionName      = "Dental Clinic"
    ds.Manufacturer         = "Carestream"

    pixels = make_panoramic(patient["scenario"])
    ds.PixelData = pixels.tobytes()

    fname = os.path.join(OUTPUT_DIR, f"{patient['id']}_panoramic.dcm")
    pydicom.dcmwrite(fname, ds)
    print(f"  Created: {fname}  (scenario: {patient['scenario']})")


if __name__ == "__main__":
    print("Generating test DICOM files…")
    for i, p in enumerate(PATIENTS):
        write_dicom(p, i)
    print(f"\nDone — {len(PATIENTS)} files in ./{OUTPUT_DIR}/")
