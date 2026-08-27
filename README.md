# Accident SOS Backend

## Getting Started

Follow these instructions to get the server running on your local machine.

### Prerequisites

- Python 3.10+ (or any recent modern Python version)

### Installation

1. Navigate to the `server` directory where the backend code resides:
   ```bash
   cd server
   ```

2. Activate your Python virtual environment (if it isn't already):
   ```powershell
   .\venv\Scripts\activate
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running the Server

To allow external hardware (on the same Wi-Fi/local network) to communicate with this server, start it with the following command:

```bash
uvicorn app:app --port 8000
```
*Note: Depending on your firewall settings, Windows may prompt you to allow Python to communicate on private networks. Make sure to click "Allow".*

---

## API Endpoints & Testing the Connection

Once the server is running, you can interact with it using the following endpoints. 

**Important:** If you are testing from an external hardware device, replace `localhost` with your computer's local IP address (e.g., `192.168.1.100`). You can find this by opening a new PowerShell window and running `ipconfig` (look for the "IPv4 Address").

### 1. Health Check
Use this endpoint to quickly test if the server is up from your browser or API testing tool (like Postman or cURL).

- **URL:** `http://localhost:8000/api/health`
- **Method:** `GET`
- **Response:**
  ```json
  {
      "status": "ok",
      "message": "Server is up and running healthy"
  }
  ```

### 2. Send Sensor Data
Use this endpoint in your hardware code to post telemetry and sensor readings to the server.

- **URL:** `http://<YOUR_COMPUTER_IP>:8000/api/sensor`
- **Method:** `POST`
- **Headers:** `Content-Type: application/json`
- **Example Body:**
  ```json
  {
      "accelerometer": {"x": 1.02, "y": -0.45, "z": 9.81},
      "gyroscope": {"x": 0.01, "y": -0.02, "z": 0.00},
      "device_id": "esp32-sensor-01"
  }
  ```
- **Response (Success):**
  ```json
  {
      "message": "Sensor data received successfully"
  }
  ```
- **Response (Error):**
  ```json
  {
      "status": "error",
      "message": "Failed to parse sensor data"
  }
  ```
  
When data is successfully sent, the server will also print the received JSON object directly into the terminal where `uvicorn` is running. This makes it incredibly easy to debug and verify the exact data your hardware is transmitting!
