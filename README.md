# Nevis Vessel Command Center

## Run

```powershell
python -m pip install -r requirements.txt
python -m streamlit run dashboard.py
```

## Application mode

Mode and the local test-data folder are configured in `nevis_api.ini`:

```ini
[app]
mode = api_test
data_dir = data
unit_api_sample_file = data/unit_api_response.xml
vessel_api_sample_file = data/vessel_visit_api_response.xml
```

Use `mode = api_test` for the saved Unit and Vessel Visit API XML responses.
Use `mode = test` for the Excel files stored in the project's `data` folder.
Use `mode = live` for the Nevis APIs.

## Live API configuration

Edit `nevis_api.ini` and enter the Nevis Basic Auth username and password:

```ini
[auth]
username = YOUR_USERNAME
password = YOUR_PASSWORD
```

The supplied configuration starts in `api_test` mode. Change `[app] mode` to
`live` after the API request has been validated.

- Vessel visits are fetched from the `vesselVisits` query.
- Only units linked to the returned working visits are fetched from `units_1`.
- Unit queries run one at a time to protect the Nevis load balancer.
- Automatic HTTP retries are disabled.
- Successful responses are cached for five minutes.
- After one API failure, a five-minute circuit breaker prevents repeated calls.
- A local safety snapshot is used temporarily if Nevis becomes unavailable.
- If no live snapshot exists yet, the Excel test snapshot keeps the interface open.

Configuration values can be overridden with environment variables such as
`NEVIS_API_USERNAME`, `NEVIS_API_PASSWORD`, `NEVIS_UNIT_API_URL`, and
`NEVIS_VESSEL_VISIT_API_URL`.

The `[app]` section of `nevis_api.ini` is authoritative for mode and test-data
location, so no command-line mode variable is required.
