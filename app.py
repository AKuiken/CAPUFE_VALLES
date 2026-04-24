# ─────────────────────────────────────────────────────────────────────────────
# CAPUFE – Servidor web  (Flask)
# ─────────────────────────────────────────────────────────────────────────────

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
    # HOY 
    df_file  = request.files.get("df_file")
    cci_file = request.files.get("cci_file")

    # Día anterior 
    df_ant_file  = request.files.get("df_file_2")
    cci_ant_file = request.files.get("cci_file_2")

    # Día posterior 
    df_pos_file  = request.files.get("df_file_3")
    cci_pos_file = request.files.get("cci_file_3")

    if not df_file or not cci_file:
        return jsonify({"error": "Se requieren ambos archivos (CLR/DF y CCI) del día HOY."}), 400

    # Si suben uno del día anterior, deben subir el par completo
    if (df_ant_file and not cci_ant_file) or (cci_ant_file and not df_ant_file):
        return jsonify({"error": "Para 'Día anterior' debes subir ambos: CLR anterior y CCI anterior."}), 400

    # Si suben uno del día posterior, deben subir el par completo
    if (df_pos_file and not cci_pos_file) or (cci_pos_file and not df_pos_file):
        return jsonify({"error": "Para 'Día posterior' debes subir ambos: CLR posterior y CCI posterior."}), 400

    try:

        clr_df = read_clr(df_file.read())
        cci_df = read_cci(cci_file.read())
        merged_hoy = build_indicators(clr_df, cci_df)

        merged_ant = None
        if df_ant_file and cci_ant_file:
            clr_ant = read_clr(df_ant_file.read())
            cci_ant = read_cci(cci_ant_file.read())
            merged_ant = build_indicators(clr_ant, cci_ant)

        merged_pos = None
        if df_pos_file and cci_pos_file:
            clr_pos = read_clr(df_pos_file.read())
            cci_pos = read_cci(cci_pos_file.read())
            merged_pos = build_indicators(clr_pos, cci_pos)

        buf = generate_excel(merged_hoy, merged_ant=merged_ant, merged_pos=merged_pos)

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
    print(f"  Local :  http://127.0.0.1:8080")
    print(f"  Red   :  http://{ip}:8080")
    print("=" * 54 + "\n")
    app.run(host="0.0.0.0", port=8080, debug=False)
     