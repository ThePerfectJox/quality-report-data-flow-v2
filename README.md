# quality-report-data-flow-v2

Quality IPC Report Data Flow with TCP

A client device asks a server device for the quality report over TLS. The server reads
the new samples from MySQL and sends them back as JSON.

## Install

Server device:

```powershell
python -m pip install -r requirements.txt
```

Client device, which only needs these packages:

```powershell
python -m pip install python-dotenv==1.2.4 openpyxl==3.1.5 tzdata==2026.3
```

`openpyxl` and `tzdata` are only needed to append the report to the Excel database.

## Run

Double-click `run.bat` in the project folder and choose:

1. Run server, on the server device
2. Run client, on the client device
3. Run config editor, on the client device

It uses `.venv` if the project has one, otherwise Python on PATH. To stop the server or the
editor, press Ctrl+C, then N at "Terminate batch job" to go back to the menu.

From a terminal instead:

```powershell
python src/run_server.py
python src/run_client.py [server-ip] [message]
python src/run_config_editor.py
```

On its first run, the server creates its TLS certificate in `certs/`.

## Append the report to the Excel master data

`run_client.py` appends the server's response to the Excel master data as soon as it receives it.
To append a saved report JSON file instead, run:

```powershell
python src/run_append_to_database.py [report.json]
```

All of this runs on the client device. The code is in `src/client/`:

| File | What it does |
| --- | --- |
| `append_to_database.py` | Runs the steps below in order |
| `report_conversion.py` | Turns each row of the `summary_report` view (`set-up/view.sql`) into a record, with the averages, Group, Brand and the server's Plant, Lokasi and Bagian |
| `inspec.py` | The in-spec calculation, using the targets and tolerances of the sample's brand |
| `master_data_mapping.py` | The dictionaries that say which view column or calculated value goes into which master data column |
| `excel_database.py` | Appends the records to the master data Excel Table and saves it safely |
| `archive.py` | Writes the archive workbook |

It needs these settings:

| Where | Setting | Purpose |
| --- | --- | --- |
| `config/client-config.json` | `MASTER_DATA.EXCEL_PATH` | The workbook holding the master data Excel Table. It must already exist |
| `config/client-config.json` | `MASTER_DATA.SHEET_NAME`, `MASTER_DATA.TABLE_NAME` | The worksheet and table. Needed only when the workbook has more than one table |
| `config/client-config.json` | `INSPEC_REQUIREMENTS` | The targets and tolerances of each brand |
| `config/server-config.json` (server device) | `PLANT`, `LOKASI`, `BAGIAN` | Where the server device is. The server sends them with each report, and they're written to every row |
| `.env` | `REPORT_JSON_PATH` | The report JSON file. Only used by `run_append_to_database.py` when it's given no file |

Copy `config/client-config.json.example` to `config/client-config.json`, then fill it in with
the config editor below or by hand. Relative paths are taken from the project folder. Close
the workbook in Excel before running.

### Config editor

To change the client settings in the browser instead of editing the JSON, choose 3 in
`run.bat`, or run:

```powershell
python src/run_config_editor.py
```

It opens `config-editor.html` at `http://127.0.0.1:8765/`. The page has a section for the
master data workbook, with a **Browse...** button, and a grid for the inspec requirements. Each
section has its own Save button and only changes its own part of `config/client-config.json`.
A save is refused with a message if a value is invalid, using the same checks as the append.
Keep the terminal window open while you edit, and press Ctrl+C to close the editor.

The editor only accepts connections from this computer. Opening `config-editor.html` directly
from the folder doesn't work, because a browser page can't save files on its own.

Where each row's values come from:

| Master data column | Source |
| --- | --- |
| `Plant/Reg`, `Lokasi`, `Bagian` | `config/server-config.json` on the server |
| `Group/Cell` | The view's `ID Group / Cell` |
| `Brand` | The view's `Nama Group / Cell`, up to the first `-`: `SAI-2` gives `SAI` |
| `Set Sample` | The view's `Set Sample`. A report with an empty Set Sample is rejected |

### Inspec requirements

Each brand in `INSPEC_REQUIREMENTS` needs all six numbers:

```json
"INSPEC_REQUIREMENTS": {
    "SAI": {
        "TARGET_CW": 2.02, "TOLERANCE_CW": 0.14,
        "TARGET_ME": 8.0,  "TOLERANCE_ME": 0.25,
        "TARGET_BE": 10.0, "TOLERANCE_BE": 0.25
    }
}
```

A measurement is in spec when `target - tolerance <= value <= target + tolerance`. Each sample
uses the entry of its brand, matched ignoring case. If a sample's brand has no entry, nothing
from the report is appended, and the error names the brand to add.

### Columns

`src/client/master_data_mapping.py` lists every master data column that is written:

- `REPORT_TO_MASTER_COLUMNS`: view columns copied as they are, such as `Loose / Void` to `VFI NC LV`.
- `CALCULATED_TO_MASTER_COLUMNS`: calculated and metadata values, such as `avg_cw` to `Avg. CW`.
  `Total VFI Cig.` is the sum of all the defect counts in `REPORT_TO_MASTER_COLUMNS`.

Other master data columns stay empty. Every view column is kept in the archive.

### Response IDs

Each row gets its `ID tanggapan` while the response is converted, before the master data is
opened. An ID is the submission date and a counter, `YYMMDD001`, such as `261106001`. The counter
continues after the last ID in `last-response.json` in the project folder, and starts at 001
on a new date. After the rows are saved, the file holds the last ID appended:

```json
{
    "last_response_id": "261106003",
    "saved_at": "2026-11-06T15:16:12+07:00"
}
```

The file is created by the first append. If the master data already has IDs from before,
create it first with the last `ID tanggapan` in the master data. If an ID is already in the
master data, nothing is appended, and the error asks you to update the file.

Each run also writes an archive next to the workbook for each Set Sample in the report, at
`archived-<PLANT>/<ISO year>/<ISO week>/<yy_mm_dd_set-sample_counter>.xlsx`, for example
`archived-RK1/2026/45/26_11_06_2_001.xlsx`. The date is the submission date, and the Set
Sample is the one of the rows in the file. The counter starts at 001 for each date and
Set Sample, and goes up by one with each archive. A report whose samples are
already in the table is rejected, so running the script twice doesn't add rows twice.

Run the tests with:

```powershell
python -m unittest discover -s tests -v
```

## Network settings

Set these in the `.env` file on each device. `.env.example` is a template.

| Setting | Device | Purpose | If empty |
| --- | --- | --- | --- |
| `SERVER_LISTEN_HOST` | Server | The IP the server listens on | Listens on all network interfaces (`0.0.0.0`) |
| `CLIENT_DEFAULT_SERVER_HOST` | Client | The server's IP. A server IP on the command line takes priority | Uses the default in `src/shared/settings.py` |
| `CLIENT_BIND_IP` | Client | This device's own Ethernet IP, used as the connection's source address | Windows chooses the network |

A real environment variable with the same name overrides the `.env` file.

## Direct Ethernet setup on Windows

Use this when two Windows devices are joined directly by an Ethernet cable.

Windows can send traffic for a `169.254.x.x` address through Wi-Fi or a VPN instead of
the cable, and the client then fails with `timed out`. Setting `CLIENT_BIND_IP` makes the
client connect from its Ethernet address.

### 1. Find each device's Ethernet IP

Run this on each device:

```powershell
ipconfig
```

Find the section for the cable connection, usually `Ethernet adapter Ethernet`, and note
its `IPv4 Address`. Don't use the addresses listed under Wi-Fi or VPN adapters.

### 2. Create `.env` on each device

The examples below use a server at `169.254.50.1` and a client at `169.254.183.0`.
Use your own addresses.

Server device:

```ini
SERVER_LISTEN_HOST=169.254.50.1
```

The server's `.env` also holds the `DB_` settings. Leave `SERVER_LISTEN_HOST` empty to listen on
all network interfaces, including Wi-Fi. Setting it limits the server to the cable, but the
server then won't start if that IP changes, and can't be reached through `127.0.0.1`.

Client device:

```ini
CLIENT_DEFAULT_SERVER_HOST=169.254.50.1
CLIENT_BIND_IP=169.254.183.0
```

Don't copy the server's `.env` to the client, because it holds the database settings.

### 3. Copy the certificate

Start the server once so it creates `certs/server.crt`. Copy that one file to `certs/` on the
client device. Never copy `certs/server.key`.

The client verifies the certificate against the fixed name `quality-report-server`, not an IP
address, so the certificate stays valid if either device's IP changes.

### 4. Test

Check the cable first, from the client device, using the example addresses:

```powershell
ping -S 169.254.183.0 -n 20 169.254.50.1
```

Then start the server, and run the client:

```powershell
python src/run_client.py
```

The client should print `Connected securely to 169.254.50.1:5000 from 169.254.183.0`, and the
server should print `Client connected from 169.254.183.0`.

### Troubleshooting

| Message | What to check |
| --- | --- |
| `CLIENT_BIND_IP is ..., but that IP can't be used on this device` | The client's Ethernet IP has changed or the cable is unplugged. Run `ipconfig` and update `.env` |
| `Could not listen on ...` | The server's Ethernet IP has changed, the cable is unplugged, or the port is in use. Run `ipconfig` and update `SERVER_LISTEN_HOST` |
| `attempted to an unreachable network` | `CLIENT_BIND_IP` and the server's IP aren't on the same network. Check `CLIENT_BIND_IP` is the Ethernet IP, not the Wi-Fi or VPN one |
| `timed out` | Check `CLIENT_BIND_IP` is set to the Ethernet IP, not the Wi-Fi or VPN one. Repeat the `ping` test, and check Windows Firewall allows inbound TCP port 5000 on the server |
| `certificate verify failed` | The client's `certs/server.crt` isn't the server's current certificate. Copy it again |

### Keeping the addresses stable

`169.254.x.x` addresses are assigned automatically and can change. For a permanent link, give
each Ethernet adapter a fixed IP on a private network that Wi-Fi and the VPN don't use, such as
`192.168.50.1` and `192.168.50.2` with subnet mask `255.255.255.0`. Check with your network
administrator that this range is free, then set each address under **Settings > Network & internet
> Ethernet > IP assignment**. Update `.env` on both devices afterwards.
