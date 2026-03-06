import socket
import traceback
from datetime import datetime

from flask import Flask, render_template, request, send_file, jsonify

from processor import read_clr, read_cci, build_indicators
from excel_builder import generate_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB


def _local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# ── Rutas ─────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", ip=_local_ip(), year=datetime.now().year)


@app.route("/process", methods=["POST"])
def process():
    df_file  = request.files.get("df_file")
    cci_file = request.files.get("cci_file")

    if not df_file or not cci_file:
        return jsonify({"error": "Se requieren ambos archivos (CLR/DF y CCI)."}), 400

    try:
        clr_df = read_clr(df_file.read())
        cci_df = read_cci(cci_file.read())
        merged = build_indicators(clr_df, cci_df)
        buf    = generate_excel(merged)

        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"Conciliacion_CLR_CCI_{ts}.xlsx"

        return send_file(
            buf,
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


# ── Inicio ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ip = _local_ip()
    print("\n" + "=" * 54)
    print("  CAPUFE – Conciliación CLR-CCI")
    print("=" * 54)
    print(f"  Local :  http://127.0.0.1:5000")
    print(f"  Red   :  http://{ip}:5000")
    print("=" * 54 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
