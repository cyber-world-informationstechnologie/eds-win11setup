## Concurrency & Rate Limiting

- The web service only processes one ISO creation request at a time.
- If a new request arrives while another is running, it responds with HTTP 429 (Too Many Requests) and the message:
  
  `Server is busy processing another image. Please wait and try again.`
- Please wait for your callback notification before submitting another request.

## Configuration

All settings are managed via `WebService.config.yaml`:

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