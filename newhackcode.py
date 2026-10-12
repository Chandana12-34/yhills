from datetime import datetime
from pathlib import Path


def fetch_crds_health(file_path):
    components = []
    error = None

    file_path = Path(file_path)

    if not file_path.is_absolute():
        file_path = Path(__file__).resolve().parent / file_path

    try:
        with file_path.open("r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()

                if not line or "|" not in line:
                    continue

                name, status = line.rsplit("|", 1)
                name = name.strip()
                status = status.strip().upper()

                if name and status in ("UP", "DOWN"):
                    components.append({
                        "name": name,
                        "status": status
                    })

        if not components:
            overall_status = "UNKNOWN"
            error = "No valid component statuses found in CRDS file"
        elif any(c["status"] == "DOWN" for c in components):
            overall_status = "DOWN"
        else:
            overall_status = "UP"

    except OSError as exc:
        overall_status = "UNKNOWN"
        error = f"Unable to read CRDS file: {exc}"

    return {
        "endpoint_name": "CRDS",
        "url": str(file_path),
        "status": overall_status,
        "details": {},
        "components": components,
        "timestamp": datetime.now().isoformat(),
        "error": error,
        "response_time": None
    }