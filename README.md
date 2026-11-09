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

```powershell
python src/run_server.py
python src/run_client.py [server-ip] [message]
```

On its first run, the server creates its TLS certificate in `certs/`.

## Append the report to the Excel database

`run_client.py` appends the server's response to the Excel database as soon as it receives it.
To append a saved report JSON file instead, run:

```powershell
python src/append_data_to_excel_database.py
```

It needs these settings:

| Where | Setting | Purpose |
| --- | --- | --- |
| `.env` | `REPORT_JSON_PATH` | The report JSON file. Only needed by `append_data_to_excel_database.py` on its own |
| `.env` | `EXCEL_DATABASE_PATH` | The workbook holding the database Excel Table. It must already exist |
| `.env` | `EXCEL_DATABASE_SHEET`, `EXCEL_DATABASE_TABLE` | The worksheet and table. Needed only when the workbook has more than one table |
| `client-config.json` | `PLANT`, `LOKASI`, `GROUP`, `BRAND`, `BAGIAN` | This device's metadata, written to every row |
| `client-config.json` | `SET_SAMPLE_THRESHOLD` | `HH:MM` in Asia/Jakarta time. A report submitted after it is Set Sample 2, otherwise 1. Exactly at the threshold is 1 |

Copy `client-config.json.example` to `client-config.json` and fill in the device's values.
Relative paths are taken from the project folder. Close the workbook in Excel before running.

Each run also writes an archive next to the workbook, at
`archived-<PLANT>/<ISO year>/<ISO week>/<yy_mm_dd_set-sample_counter>.xlsx`, for example
`archived-RK1/2026/45/26_11_06_2_001.xlsx`. The counter starts at 001 for each date and
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
