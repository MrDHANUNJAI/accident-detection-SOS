from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse

from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook

import json
import urllib.parse


# ============================================================
# APPLICATION
# ============================================================

app = FastAPI(
    title="Smart Vehicle SOS System",
    version="2.0.0"
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)


# ============================================================
# BASE DIRECTORIES
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"

DATA_DIR.mkdir(
    exist_ok=True
)


# ============================================================
# FILE PATHS
# ============================================================

INDEX_FILE = BASE_DIR / "index.html"

JSON_FILE = DATA_DIR / "sos_events.json"

EXCEL_FILE = DATA_DIR / "sos_events.xlsx"

SENSOR_FILE = DATA_DIR / "sensor_data.json"

SETTINGS_FILE = DATA_DIR / "settings.json"


# ============================================================
# FALLBACK GPS LOCATION
# ============================================================

FALLBACK_LATITUDE = 13.629000

FALLBACK_LONGITUDE = 78.485000


# ============================================================
# ROOT DASHBOARD
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
def dashboard():

    if not INDEX_FILE.exists():

        return HTMLResponse(
            content="""
            <!DOCTYPE html>
            <html>
            <head>
                <title>SOS Dashboard Error</title>
            </head>
            <body>
                <h1>Dashboard not found</h1>
                <p>
                    Please place index.html in the same
                    folder as app.py.
                </p>
            </body>
            </html>
            """,
            status_code=500
        )

    return HTMLResponse(
        content=INDEX_FILE.read_text(
            encoding="utf-8"
        )
    )


# ============================================================
# LOAD EVENTS
# ============================================================

def load_events():

    if not JSON_FILE.exists():

        return []

    try:

        with open(
            JSON_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            data = json.load(file)

            if isinstance(data, list):

                return data

            return []

    except Exception as error:

        print(
            "JSON LOAD ERROR:",
            error
        )

        return []


# ============================================================
# SAVE EVENTS
# ============================================================

def save_events(events):

    with open(
        JSON_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            events,
            file,
            indent=4
        )


# ============================================================
# GENERATE EVENT ID
# ============================================================

def generate_event_id():

    events = load_events()

    highest = 0

    for event in events:

        event_id = str(
            event.get(
                "eventId",
                ""
            )
        )

        if event_id.startswith("EVT-"):

            try:

                number = int(
                    event_id.replace(
                        "EVT-",
                        ""
                    )
                )

                if number > highest:

                    highest = number

            except ValueError:

                pass

    return f"EVT-{highest + 1:06d}"


# ============================================================
# CREATE EXCEL
# ============================================================

def create_excel():

    if EXCEL_FILE.exists():

        return

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "SOS Events"

    headers = [

        "Event ID",
        "Vehicle ID",
        "Driver",
        "Alert Type",

        "Acceleration X",
        "Acceleration Y",
        "Acceleration Z",
        "Acceleration Magnitude",

        "Gyroscope X",
        "Gyroscope Y",
        "Gyroscope Z",

        "Tilt",

        "Latitude",
        "Longitude",

        "GPS Detected",
        "Location Source",

        "Sensor Source",

        "WiFi Status",
        "WhatsApp Status",

        "Event Status",

        "Cancellation Time",
        "Response Time",

        "Date/Time"
    ]

    sheet.append(headers)

    for cell in sheet[1]:

        cell.font = cell.font.copy(
            bold=True
        )

    for column in sheet.columns:

        column_letter = (
            column[0].column_letter
        )

        sheet.column_dimensions[
            column_letter
        ].width = 20

    workbook.save(
        EXCEL_FILE
    )


# ============================================================
# SAVE EVENT TO EXCEL
# ============================================================

def save_event_to_excel(event):

    create_excel()

    workbook = load_workbook(
        EXCEL_FILE
    )

    sheet = workbook[
        "SOS Events"
    ]

    row = [

        event.get(
            "eventId",
            ""
        ),

        event.get(
            "vehicleId",
            ""
        ),

        event.get(
            "driver",
            ""
        ),

        event.get(
            "alertType",
            ""
        ),

        event.get(
            "accelerationX",
            0
        ),

        event.get(
            "accelerationY",
            0
        ),

        event.get(
            "accelerationZ",
            0
        ),

        event.get(
            "accelerationMagnitude",
            0
        ),

        event.get(
            "gyroX",
            0
        ),

        event.get(
            "gyroY",
            0
        ),

        event.get(
            "gyroZ",
            0
        ),

        event.get(
            "tilt",
            0
        ),

        event.get(
            "latitude",
            FALLBACK_LATITUDE
        ),

        event.get(
            "longitude",
            FALLBACK_LONGITUDE
        ),

        event.get(
            "gpsDetected",
            False
        ),

        event.get(
            "locationSource",
            "FALLBACK"
        ),

        event.get(
            "sensorSource",
            "UNKNOWN"
        ),

        event.get(
            "wifiStatus",
            "UNKNOWN"
        ),

        event.get(
            "whatsappStatus",
            "NOT_CONFIGURED"
        ),

        event.get(
            "eventStatus",
            "ACTIVE"
        ),

        event.get(
            "cancellationTime",
            ""
        ),

        event.get(
            "responseTime",
            ""
        ),

        event.get(
            "dateTime",
            ""
        )
    ]

    sheet.append(row)

    workbook.save(
        EXCEL_FILE
    )


# ============================================================
# UPDATE EVENT IN EXCEL
# ============================================================

def update_event_in_excel(event):

    create_excel()

    workbook = load_workbook(
        EXCEL_FILE
    )

    sheet = workbook[
        "SOS Events"
    ]

    event_id = event.get(
        "eventId"
    )

    for row in sheet.iter_rows(
        min_row=2
    ):

        if row[0].value == event_id:

            values = [

                event.get(
                    "eventId",
                    ""
                ),

                event.get(
                    "vehicleId",
                    ""
                ),

                event.get(
                    "driver",
                    ""
                ),

                event.get(
                    "alertType",
                    ""
                ),

                event.get(
                    "accelerationX",
                    0
                ),

                event.get(
                    "accelerationY",
                    0
                ),

                event.get(
                    "accelerationZ",
                    0
                ),

                event.get(
                    "accelerationMagnitude",
                    0
                ),

                event.get(
                    "gyroX",
                    0
                ),

                event.get(
                    "gyroY",
                    0
                ),

                event.get(
                    "gyroZ",
                    0
                ),

                event.get(
                    "tilt",
                    0
                ),

                event.get(
                    "latitude",
                    FALLBACK_LATITUDE
                ),

                event.get(
                    "longitude",
                    FALLBACK_LONGITUDE
                ),

                event.get(
                    "gpsDetected",
                    False
                ),

                event.get(
                    "locationSource",
                    "FALLBACK"
                ),

                event.get(
                    "sensorSource",
                    "UNKNOWN"
                ),

                event.get(
                    "wifiStatus",
                    "UNKNOWN"
                ),

                event.get(
                    "whatsappStatus",
                    "NOT_CONFIGURED"
                ),

                event.get(
                    "eventStatus",
                    "ACTIVE"
                ),

                event.get(
                    "cancellationTime",
                    ""
                ),

                event.get(
                    "responseTime",
                    ""
                ),

                event.get(
                    "dateTime",
                    ""
                )
            ]

            for index, value in enumerate(
                values,
                start=1
            ):

                sheet.cell(
                    row=row[0].row,
                    column=index
                ).value = value

            break

    workbook.save(
        EXCEL_FILE
    )


# ============================================================
# HEALTH API
# ============================================================

@app.get(
    "/api/health"
)
def health():

    return {

        "status":
            "ok",

        "message":
            "Smart Vehicle SOS Server is running",

        "database":
            "Excel + JSON",

        "excelExists":
            EXCEL_FILE.exists(),

        "totalEvents":
            len(load_events())
    }


# ============================================================
# GET ALL EVENTS
# ============================================================

@app.get(
    "/api/events"
)
def get_events():

    events = load_events()

    events = list(
        reversed(events)
    )

    return {

        "status":
            "success",

        "count":
            len(events),

        "events":
            events
    }


# ============================================================
# RECEIVE SOS
# ============================================================

@app.post(
    "/api/sos"
)
async def receive_sos(
    request: Request
):

    try:

        data = await request.json()

        print()
        print(
            "=========================================="
        )

        print(
            "       NEW SOS ALERT RECEIVED"
        )

        print(
            "=========================================="
        )

        print(
            "DATA:",
            data
        )


        # ====================================================
        # EVENT ID
        # ====================================================

        event_id = generate_event_id()


        # ====================================================
        # DATE / TIME
        # ====================================================

        date_time = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # ====================================================
        # GPS
        # ====================================================

        gps_detected = bool(
            data.get(
                "gpsDetected",
                False
            )
        )

        if gps_detected:

            latitude = data.get(
                "latitude",
                FALLBACK_LATITUDE
            )

            longitude = data.get(
                "longitude",
                FALLBACK_LONGITUDE
            )

            location_source = "GPS"

        else:

            latitude = FALLBACK_LATITUDE

            longitude = FALLBACK_LONGITUDE

            location_source = "FALLBACK"


        # ====================================================
        # CREATE EVENT
        # ====================================================

        event = {

            "eventId":
                event_id,

            "vehicleId":
                data.get(
                    "vehicleId",
                    "VEHICLE-001"
                ),

            "driver":
                data.get(
                    "driver",
                    "UNKNOWN"
                ),

            "alertType":
                data.get(
                    "alertType",
                    data.get(
                        "event",
                        "MANUAL_SOS"
                    )
                ),

            "accelerationX":
                data.get(
                    "accelerationX",
                    0
                ),

            "accelerationY":
                data.get(
                    "accelerationY",
                    0
                ),

            "accelerationZ":
                data.get(
                    "accelerationZ",
                    0
                ),

            "accelerationMagnitude":
                data.get(
                    "accelerationMagnitude",
                    0
                ),

            "gyroX":
                data.get(
                    "gyroX",
                    0
                ),

            "gyroY":
                data.get(
                    "gyroY",
                    0
                ),

            "gyroZ":
                data.get(
                    "gyroZ",
                    0
                ),

            "tilt":
                data.get(
                    "tilt",
                    0
                ),

            "latitude":
                latitude,

            "longitude":
                longitude,

            "gpsDetected":
                gps_detected,

            "locationSource":
                location_source,

            "sensorSource":
                data.get(
                    "sensorSource",
                    "UNKNOWN"
                ),

            "wifiStatus":
                data.get(
                    "wifiStatus",
                    "CONNECTED"
                ),

            "whatsappStatus":
                data.get(
                    "whatsappStatus",
                    "PENDING"
                ),

            "eventStatus":
                "ACTIVE",

            "cancellationTime":
                "",

            "responseTime":
                "",

            "dateTime":
                date_time
        }


        # ====================================================
        # SAVE JSON
        # ====================================================

        events = load_events()

        events.append(
            event
        )

        save_events(
            events
        )


        # ====================================================
        # SAVE EXCEL
        # ====================================================

        save_event_to_excel(
            event
        )


        # ====================================================
        # CONSOLE
        # ====================================================

        print(
            "EVENT SAVED:",
            event_id
        )

        print(
            "LOCATION:",
            latitude,
            longitude
        )

        print(
            "STATUS:",
            "ACTIVE"
        )

        print(
            "=========================================="
        )


        # ====================================================
        # RESPONSE
        # ====================================================

        return {

            "status":
                "success",

            "message":
                "SOS received and stored successfully",

            "event":
                event
        }


    except Exception as error:

        print(
            "SOS ERROR:",
            error
        )

        return {

            "status":
                "error",

            "message":
                str(error)
        }


# ============================================================
# SENSOR DATA
# ============================================================

@app.post(
    "/api/sensor"
)
async def receive_sensor(
    request: Request
):

    try:

        data = await request.json()

        existing = []

        if SENSOR_FILE.exists():

            try:

                with open(
                    SENSOR_FILE,
                    "r",
                    encoding="utf-8"
                ) as file:

                    existing = json.load(
                        file
                    )

                    if not isinstance(
                        existing,
                        list
                    ):

                        existing = []

            except Exception:

                existing = []


        existing.append(
            data
        )


        with open(
            SENSOR_FILE,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                existing,
                file,
                indent=4
            )


        return {

            "status":
                "success",

            "message":
                "Sensor data received",

            "data":
                data
        }


    except Exception as error:

        return {

            "status":
                "error",

            "message":
                str(error)
        }


# ============================================================
# GET SENSOR DATA
# ============================================================

@app.get(
    "/api/sensor"
)
def get_sensor():

    if not SENSOR_FILE.exists():

        return []


    try:

        with open(
            SENSOR_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            return json.load(
                file
            )


    except Exception:

        return []


# ============================================================
# GET SINGLE EVENT
# ============================================================

@app.get(
    "/api/events/{event_id}"
)
def get_single_event(
    event_id: str
):

    events = load_events()

    for event in events:

        if event.get(
            "eventId"
        ) == event_id:

            return {

                "status":
                    "success",

                "event":
                    event
            }


    return {

        "status":
            "error",

        "message":
            "Event not found"
    }


# ============================================================
# CANCEL EVENT
# ============================================================

@app.patch(
    "/api/events/{event_id}/cancel"
)
def cancel_event(
    event_id: str
):

    events = load_events()

    for event in events:

        if event.get(
            "eventId"
        ) == event_id:


            event[
                "eventStatus"
            ] = "CANCELLED"


            event[
                "cancellationTime"
            ] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )


            event[
                "responseTime"
            ] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )


            save_events(
                events
            )


            update_event_in_excel(
                event
            )


            return {

                "status":
                    "success",

                "message":
                    "SOS cancelled",

                "event":
                    event
            }


    return {

        "status":
            "error",

        "message":
            "Event not found"
    }


# ============================================================
# RESOLVE EVENT
# ============================================================

@app.patch(
    "/api/events/{event_id}/resolve"
)
def resolve_event(
    event_id: str
):

    events = load_events()

    for event in events:

        if event.get(
            "eventId"
        ) == event_id:


            event[
                "eventStatus"
            ] = "RESOLVED"


            event[
                "responseTime"
            ] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )


            save_events(
                events
            )


            update_event_in_excel(
                event
            )


            return {

                "status":
                    "success",

                "message":
                    "SOS resolved",

                "event":
                    event
            }


    return {

        "status":
            "error",

        "message":
            "Event not found"
    }


# ============================================================
# SETTINGS
# ============================================================

def load_settings():

    if not SETTINGS_FILE.exists():

        return {

            "vehicleId":
                "VEHICLE-001",

            "driver":
                "Driver-01",

            "whatsappPhone":
                ""
        }


    try:

        with open(
            SETTINGS_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            return json.load(
                file
            )


    except Exception:

        return {

            "vehicleId":
                "VEHICLE-001",

            "driver":
                "Driver-01",

            "whatsappPhone":
                ""
        }


def save_settings(
    settings
):

    with open(
        SETTINGS_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            settings,
            file,
            indent=4
        )


# ============================================================
# GET SETTINGS
# ============================================================

@app.get(
    "/api/settings"
)
def get_settings():

    return load_settings()


# ============================================================
# UPDATE SETTINGS
# ============================================================

@app.put(
    "/api/settings"
)
async def update_settings(
    request: Request
):

    data = await request.json()

    settings = load_settings()

    settings.update(
        data
    )

    save_settings(
        settings
    )

    return {

        "status":
            "success",

        "settings":
            settings
    }


# ============================================================
# WHATSAPP MESSAGE
# ============================================================

@app.get(
    "/api/whatsapp/{event_id}"
)
def whatsapp_event(
    event_id: str
):

    events = load_events()

    settings = load_settings()

    phone = settings.get(
        "whatsappPhone",
        ""
    )


    for event in events:

        if event.get(
            "eventId"
        ) == event_id:


            message = (
                "SMART VEHICLE SOS ALERT\n\n"

                f"Event ID: "
                f"{event.get('eventId')}\n"

                f"Vehicle: "
                f"{event.get('vehicleId')}\n"

                f"Driver: "
                f"{event.get('driver')}\n"

                f"Alert Type: "
                f"{event.get('alertType')}\n"

                f"Status: "
                f"{event.get('eventStatus')}\n\n"

                f"Location: "
                f"{event.get('latitude')}, "
                f"{event.get('longitude')}\n"

                f"Time: "
                f"{event.get('dateTime')}"
            )


            encoded_message = (
                urllib.parse.quote(
                    message
                )
            )


            if phone:

                url = (
                    f"https://wa.me/"
                    f"{phone}"
                    f"?text="
                    f"{encoded_message}"
                )

            else:

                url = (
                    f"https://wa.me/"
                    f"?text="
                    f"{encoded_message}"
                )


            return {

                "status":
                    "success",

                "url":
                    url
            }


    return {

        "status":
            "error",

        "message":
            "Event not found"
    }


# ============================================================
# EXPORT EXCEL
# ============================================================

@app.get(
    "/api/export/excel"
)
def export_excel():

    create_excel()

    return FileResponse(

        path=str(
            EXCEL_FILE
        ),

        filename=
            "sos_events.xlsx",

        media_type=
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
    )


# ============================================================
# STARTUP
# ============================================================

@app.on_event(
    "startup"
)
def startup():

    create_excel()


    if not JSON_FILE.exists():

        save_events([])


    print()
    print(
        "=========================================="
    )

    print(
        "       SMART VEHICLE SOS SYSTEM"
    )

    print(
        "=========================================="
    )

    print(
        "Dashboard : "
        "http://127.0.0.1:8001/"
    )

    print(
        "Health    : "
        "http://127.0.0.1:8001/api/health"
    )

    print(
        "SOS       : "
        "http://127.0.0.1:8001/api/sos"
    )

    print(
        "Events    : "
        "http://127.0.0.1:8001/api/events"
    )

    print(
        "Excel     : "
        "http://127.0.0.1:8001/api/export/excel"
    )

    print(
        "=========================================="
    )