"""The dictionaries that turn the server's report (the summary_report view) into master data rows.

Edit this file when a column of set-up/view.sql or of the master data Excel Table changes.

- REPORT_TO_MASTER_COLUMNS: view columns copied straight into a master data column.
- CALCULATED_TO_MASTER_COLUMNS: values the client calculates or adds, and their master column.
- MASTER_DATA_COLUMN_MAPPING: both together. Only these master columns are written; the
  others stay empty, or get the column's table formula if it has one.

Master columns are matched by name, ignoring case and repeated spaces, never by position.
Every view column, mapped or not, is also kept in the archive workbook.
"""

from client.values import normalize_source_key

# region VIEW COLUMNS (set-up/view.sql), as normalized by normalize_source_key()
# "Set Sample" in the view is ProcessH_Set: the sample's Set Sample. It must not be empty.
SOURCE_SET_SAMPLE_KEY = normalize_source_key("Set Sample")
# The group leader's member ID, written to Group/Cell.
SOURCE_GROUP_ID_KEY = normalize_source_key("ID Group / Cell")
# The group name, such as "SAI-2". The brand is the part before the first "-".
SOURCE_GROUP_NAME_KEY = normalize_source_key("Nama Group / Cell")
BRAND_SEPARATOR = "-"
SOURCE_WORKER_ID_KEY = normalize_source_key("ID PPSKT")
SOURCE_START_KEY = normalize_source_key("Start")
CW_SOURCE_KEYS = (normalize_source_key("Avg CW [1]"), normalize_source_key("Avg CW [2]"))
BE_SOURCE_KEYS = tuple(normalize_source_key(f"UB [{number}]") for number in range(1, 7))
ME_SOURCE_KEYS = tuple(normalize_source_key(f"UH [{number}]") for number in range(1, 7))
# endregion

# region VIEW COLUMN -> MASTER DATA COLUMN
# The visual defect (VFI NC) counts. The view's "Trial 1" and "Reserved 2" to "Reserved 10"
# have no master column, so they're only archived.
REPORT_TO_MASTER_COLUMNS = {
    "Loose / Void": "VFI NC LV",
    "Wrinkled": "VFI NC Wrinkled",
    "Improper Cut 1": "VFI NC IC 01",
    "Improper Cut 3": "VFI NC IC 03",
    "Improper Cut 4": "VFI NC IC 04",
    "Improper Cut 5": "VFI NC IC 05",
    "TPIS": "VFI NC TPIS",
    "Weak / Open 1": "VFI NC Weak/Open 01",
    "Weak / Open 3": "VFI NC Weak/Open 03",
    "Weak / Open 4": "VFI NC Weak/Open 04",
    "Weak / Open 5": "VFI NC Weak/Open 05",
    "Weak / Open 6": "VFI NC Weak/Open 06",
    "Tear / Ragged 1": "VFI NC Tear/Ragged 01",
    "Tear / Ragged 2": "VFI NC Tear/Ragged 02",
    "TBU": "VFI NC TBU",
    "Pleat": "VFI NC Pleat",
    "Tear / Hole": "VFI NC Tear/Hole",
    "Smeared / Dirty": "VFI NC Smeared/Dirty",
    "Artwork Alignment": "VFI NC Artwork Alignment",
    "Incompletely Field": "VFI NC Incompletely Filled",
    "Foreign Matter": "VFI NC Foreign Matter",
    "Clove Spot": "VFI NC Clove Spot",
}
# The same view columns, by normalized key: "Loose / Void" becomes "loose_void".
DEFECT_SOURCE_KEYS = tuple(normalize_source_key(column) for column in REPORT_TO_MASTER_COLUMNS)
# endregion

# region CALCULATED AND METADATA FIELDS
RESPONSE_ID_KEY = "tanggapan_id"
SUBMITTED_AT_KEY = "tanggal_pengiriman"
INSPECTED_AT_KEY = "tanggal_pemeriksaan"  # From the view's "Start".
WORKER_ID_KEY = "no_id_pekerja"  # From the view's "ID PPSKT".
SET_SAMPLE_KEY = SOURCE_SET_SAMPLE_KEY  # From the view's "Set Sample".
GROUP_KEY = "group"  # From the view's "ID Group / Cell".
BRAND_KEY = "brand"  # From the view's "Nama Group / Cell": "SAI-2" gives "SAI".
PLANT_KEY = "plant"  # From server-config.json on the server, sent in each response.
LOKASI_KEY = "lokasi"  # From server-config.json.
BAGIAN_KEY = "bagian"  # From server-config.json.

CALCULATED_TO_MASTER_COLUMNS = {
    RESPONSE_ID_KEY: "ID tanggapan",
    SUBMITTED_AT_KEY: "Tanggal pengiriman",
    INSPECTED_AT_KEY: "Tanggal Pemeriksaan",
    PLANT_KEY: "Plant/Reg",
    LOKASI_KEY: "Lokasi",
    GROUP_KEY: "Group/Cell",
    WORKER_ID_KEY: "No ID Pekerja",
    BRAND_KEY: "Brand",
    SET_SAMPLE_KEY: "Set Sample",
    BAGIAN_KEY: "Bagian",
    "avg_cw": "Avg. CW",
    "avg_dia_be": "Avg. Dia. BE",
    "avg_dia_me": "Avg. Dia. ME",
    "in_spec_cw": "%In Spec CW",
    "in_spec_be": "%In Spec Dia. BE",
    "in_spec_me": "%In Spec Dia. ME",
}

# Column names of the calculated and metadata fields in the archive workbook.
# View columns keep their name from the view.
ARCHIVE_FIELD_LABELS = {
    RESPONSE_ID_KEY: "Tanggapan ID",
    SUBMITTED_AT_KEY: "Tanggal Pengiriman",
    INSPECTED_AT_KEY: "Tanggal Pemeriksaan",
    PLANT_KEY: "Plant",
    LOKASI_KEY: "Lokasi",
    GROUP_KEY: "Group",
    WORKER_ID_KEY: "No Id Pekerja",
    BRAND_KEY: "Brand",
    BAGIAN_KEY: "Bagian",
    SET_SAMPLE_KEY: "Set Sample",
    "avg_cw_first": "Avg Cw [First]",
    "avg_cw_second": "Avg Cw [Second]",
    "avg_cw": "Avg Cw",
    "avg_dia_be": "Avg Dia Be",
    "avg_dia_me": "Avg Dia Me",
    "in_spec_cw": "In Spec Cw",
    "in_spec_be": "In Spec Be",
    "in_spec_me": "In Spec Me",
}
# endregion

# region THE COMPLETE MAPPING
# Normalized key -> master data column. This is what the append writes.
MASTER_DATA_COLUMN_MAPPING = {
    **CALCULATED_TO_MASTER_COLUMNS,
    **dict(zip(DEFECT_SOURCE_KEYS, REPORT_TO_MASTER_COLUMNS.values())),
}

# The append stops if the table lacks one of these columns. A mapped column that isn't
# required is skipped with a warning. Remove a key from this set to make its column optional.
REQUIRED_MASTER_DATA_KEYS = frozenset(MASTER_DATA_COLUMN_MAPPING)

# A sample without one of these view columns means the view changed, so it's rejected.
# Their values may still be empty: measurements are then left out of the averages.
REQUIRED_SOURCE_KEYS = (
    SOURCE_WORKER_ID_KEY,
    SOURCE_START_KEY,
    SOURCE_SET_SAMPLE_KEY,
    SOURCE_GROUP_ID_KEY,
    SOURCE_GROUP_NAME_KEY,
    *CW_SOURCE_KEYS,
    *BE_SOURCE_KEYS,
    *ME_SOURCE_KEYS,
    *DEFECT_SOURCE_KEYS,
)

# A sample counts as already in the master data when all of these match an existing row.
# The view has no record ID, so these fields together identify one measurement.
DUPLICATE_CHECK_KEYS = (WORKER_ID_KEY, INSPECTED_AT_KEY, "avg_cw", "avg_dia_be", "avg_dia_me")
DUPLICATE_FLOAT_DECIMALS = 6
# endregion
