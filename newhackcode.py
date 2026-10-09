from datetime import datetime


def fetch_crds_health(file_path):
    components = []
    error = None

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()

                # Skip empty lines or lines without a separator
                if not line or "|" not in line:
                    continue

                # Expected format: ComponentName | UP/DOWN
                name, status = line.rsplit("|", 1)

                name = name.strip()
                status = status.strip().upper()

                if not name:
                    continue

                # Accept only recognized component statuses
                if status not in ("UP", "DOWN"):
                    continue

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
        "url": file_path,
        "status": overall_status,
        "details": {},
        "components": components,
        "timestamp": datetime.now().isoformat(),
        "error": error,
        "response_time": None
    }