import csv
import json
import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc

APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
CONFIG_FILE = APP_DIR / "Script" / "sap_login_config.json"
if not CONFIG_FILE.exists():
    CONFIG_FILE = RESOURCE_DIR / "Script" / "sap_login_config.json"


def resolve_path(path_value, resource=False):
    path = Path(path_value)
    if path.is_absolute():
        return path
    if resource:
        app_path = APP_DIR / path
        if app_path.exists():
            return app_path
        return RESOURCE_DIR / path
    return APP_DIR / path


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def setup_logging(paths_config):
    now = datetime.now()
    log_root = resolve_path(paths_config.get("log_root", "Logs"))
    log_dir = log_root / now.strftime("%Y") / now.strftime("%m")
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"XuatDuLieuSAP_{now.strftime('%Y%m%d')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")],
    )
    logging.info("Log file: %s", log_file)
    return log_file


def log_subprocess_output(name, result):
    logging.info("%s return code: %s", name, result.returncode)
    if result.stdout:
        logging.info("%s stdout:\n%s", name, result.stdout.strip())
    if result.stderr:
        logging.warning("%s stderr:\n%s", name, result.stderr.strip())


def main():
    config = load_config()
    paths_config = config.get("paths", {})
    log_file = setup_logging(paths_config)

    logging.info("Start XuatDuLieuSAP")
    logging.info("Config file: %s", CONFIG_FILE)

    vbs_script = resolve_path(paths_config.get("vbs_script", "Script/MB52.vbs"), resource=True)
    export_folder = resolve_path(paths_config.get("export_folder", "Data"))
    export_file = paths_config.get("export_file", "MB52.XLSX")
    input_file = export_folder / export_file
    output_file = resolve_path(paths_config.get("csv_file", str(input_file.with_suffix(".csv"))))

    logging.info("VBS script: %s", vbs_script)
    logging.info("Excel input: %s", input_file)
    logging.info("CSV output: %s", output_file)

    logging.info("Run SAP export script")
    vbs_result = subprocess.run(["cscript", "//NoLogo", str(vbs_script)], capture_output=True, text=True)
    log_subprocess_output("cscript MB52.vbs", vbs_result)
    if vbs_result.returncode != 0:
        raise RuntimeError("VBS script failed")

    logging.info("Read Excel file")
    df = pd.read_excel(input_file, engine="openpyxl")
    logging.info("Excel rows: %s", len(df))

    logging.info("Write CSV file")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_file, sep=";", index=False, encoding="utf-8-sig", quoting=csv.QUOTE_NONE, escapechar="\\")

    logging.info("Clean escaped quotes in CSV")
    with open(output_file, "r", encoding="utf-8-sig") as f:
        content = f.read()
    with open(output_file, "w", encoding="utf-8-sig") as f:
        f.write(content.replace('\\"', '"'))

    logging.info("Converted: %s -> %s", input_file, output_file)

    logging.info("Connect SQL and truncate dbo.MB52")
    conn = pyodbc.connect("DRIVER={SQL Server};SERVER=10.239.1.54;DATABASE=SAPData;UID=sa;PWD=123456")
    conn.execute("TRUNCATE TABLE dbo.MB52")
    conn.commit()
    conn.close()
    logging.info("Truncated table dbo.MB52")

    logging.info("Run BCP import")
    result = subprocess.run(
        ["bcp", "SAPData.dbo.MB52", "in", str(output_file), "-c", "-t;", r"-r\n", "-F", "3", "-S", "10.239.1.54", "-U", "sa", "-P", "123456"],
        capture_output=True,
        text=True,
    )
    log_subprocess_output("bcp import", result)
    if result.returncode != 0:
        raise RuntimeError("BCP import failed")

    logging.info("Finished XuatDuLieuSAP successfully")
    return log_file


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Program failed")
        sys.exit(1)
