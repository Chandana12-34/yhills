import requests
import certifi
import ssl

print("Python SSL defaults:")
print(ssl.get_default_verify_paths())

print("\nRequests CA bundle:")
print(requests.certs.where())

print("\nCertifi CA bundle:")
print(certifi.where())

url = "https://api.llm.cib.echonet/v1/openai/chat/completions"

try:
    response = requests.get(url, timeout=15, verify=True)
    print("HTTP status:", response.status_code)
except requests.exceptions.SSLError as exc:
    print("TLS certificate error:", exc)
except requests.exceptions.RequestException as exc:
    print("Connection error:", exc)