# Windows 11 ISO Creation Web Service

This PowerShell web service allows you to remotely trigger the creation of customized Windows 11 ISO files via HTTP POST requests. It supports header authentication and sends a callback notification when the ISO is ready.

## Concurrency & Rate Limiting

- The web service only processes one ISO creation request at a time.
- If a new request arrives while another is running, it responds with HTTP 429 (Too Many Requests) and the message:

  `Server is busy processing another image. Please wait and try again.`

- Please wait for your callback notification before submitting another request.

## Configuration

All settings are managed via `automation/WebService.config.yaml`:

```yaml
auth_header: "Bearer SuperSecretToken123"
port: 8080
input_base: "C:\ISOs"
output_base: "C:\projekte\eds-win11setup\Dist\WinIso"
```

- `auth_header`: The required Authorization header value
- `port`: The port to listen on
- `input_base`: The base folder for input ISOs
- `output_base`: The base folder for output ISOs

Edit this file to change authentication, port, or folder paths. Escaping for Windows paths is handled automatically if you use double quotes.

## Features

- Simple HTTP POST endpoint
- Header authentication with a fixed token
- Accepts ISO location, cumulative update path, Windows edition, and callback URL
- Starts the installer script with provided options
- Sends a POST request to the callback URL with the ISO name when done
- All parameters can be sent in the JSON body or as query parameters

## Usage

### 1. Start the Web Service

Run as administrator:

```powershell
cd automation
./WebService.ps1
```

### 2. Send a POST Request

Example using `curl`:

```bash
curl -X POST http://localhost:8080/ \
  -H "Authorization: Bearer SuperSecretToken123" \
  -H "Content-Type: application/json" \
  -d '{
    "iso_path": "C:\\ISOs\\Win11.iso",
    "callback_url": "http://yourserver/callback",
    "cu_path": "C:\\Updates\\windows11cu.msu",
    "win_edition": "Windows 11 Pro"
  }'
```

#### Parameters

- `iso_path` (required): Path to the source Windows 11 ISO
- `callback_url` (required): URL to notify when ISO creation is done
- `cu_path` (optional): Path to cumulative update (.msu)
- `win_edition` (optional): Windows edition name (default: "Windows 11 Pro")

You can also pass parameters as query strings:

```
POST http://localhost:8080/?iso_path=...&callback_url=...&cu_path=...&win_edition=...
```

### 3. Callback Notification

When ISO creation is complete, the service sends a POST request to your callback URL:

```json
{
  "iso_name": "winiso_<random-guid>.iso"
}
```

## Authentication

- Add header: `Authorization: Bearer SuperSecretToken123`
- Change the token in `automation/WebService.config.yaml` as needed

## Requirements

- PowerShell 5.1+
- Administrator privileges
- Windows ADK (for ISO creation)
- The main installer script: `automation/Create-Windows11Installer.ps1`
- Module: `powershell-yaml` (for config parsing)

## Notes

- The service runs synchronously and will block until ISO creation is complete
- The ISO is saved in the configured output folder with a random GUID name
- Errors are returned as HTTP 401 (unauthorized), 405 (method not allowed), 429 (busy), or 500 (internal error)

## Example Response

- On success: `ISO creation started. Callback will be sent when done.`
- On error: `Error: <message>`

---

For more details, see the main project README.
