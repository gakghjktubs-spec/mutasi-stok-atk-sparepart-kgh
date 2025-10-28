from flask import Flask, render_template, request, redirect, flash, jsonify, send_file, url_for
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
app.secret_key = "stok-app-secret"

DATA_DIR = "data"
STOK_FILE = os.path.join(DATA_DIR, "stok.xlsx")
MUTASI_FILE = os.path.join(DATA_DIR, "mutasi.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

if not os.path.exists(STOK_FILE):
    pd.DataFrame(columns=["Kode Barang", "Nama Barang", "Saldo"]).to_excel(STOK_FILE, index=False)
if not os.path.exists(MUTASI_FILE):
    pd.DataFrame(columns=["Tanggal", "Kode Barang", "Nama Barang", "Jenis", "Jumlah", "Keterangan", "Nama Input"]).to_excel(MUTASI_FILE, index=False)

def load_stok():
    return pd.read_excel(STOK_FILE)

def load_mutasi():
    return pd.read_excel(MUTASI_FILE)

def save_stok(df):
    df.to_excel(STOK_FILE, index=False)

def save_mutasi(df):
    df.to_excel(MUTASI_FILE, index=False)

@app.route("/api/get_barang_list")
def api_get_barang_list():
    df = load_stok()
    data = [{"kode": row["Kode Barang"], "nama": row["Nama Barang"], "saldo": float(row["Saldo"])} for _, row in df.iterrows()]
    return jsonify(data)

@app.route("/api/get_saldo")
def api_get_saldo():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"saldo": None})
    df = load_stok()
    match = df[df["Kode Barang"] == q]
    if not match.empty:
        return jsonify({"saldo": float(match.iloc[0]["Saldo"])})
    match2 = df[df["Nama Barang"].str.contains(q, case=False, na=False)]
    if not match2.empty:
        return jsonify({"saldo": float(match2.iloc[0]["Saldo"])})
    return jsonify({"saldo": None})

@app.route("/api/add_mutasi", methods=["POST"])
def api_add_mutasi():
    data = request.json or request.form
    kode = (data.get("kode") or "").strip()
    jenis = data.get("jenis")
    jumlah = float(data.get("jumlah") or 0)
    keterangan = data.get("keterangan", "")
    nama_input = data.get("nama_input", "")
    stok_df = load_stok()
    mutasi_df = load_mutasi()
    if kode == "" or kode not in stok_df["Kode Barang"].values:
        return jsonify({"ok": False, "message": "Kode barang tidak ditemukan."}), 400
    if jenis == "masuk":
        stok_df.loc[stok_df["Kode Barang"] == kode, "Saldo"] += jumlah
    elif jenis == "keluar":
        stok_df.loc[stok_df["Kode Barang"] == kode, "Saldo"] -= jumlah
    nama_barang = stok_df.loc[stok_df["Kode Barang"] == kode, "Nama Barang"].values[0]
    new_row = {"Tanggal": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "Kode Barang": kode,
               "Nama Barang": nama_barang,
               "Jenis": jenis,
               "Jumlah": jumlah,
               "Keterangan": keterangan,
               "Nama Input": nama_input}
    mutasi_df = pd.concat([mutasi_df, pd.DataFrame([new_row])], ignore_index=True)
    save_stok(stok_df)
    save_mutasi(mutasi_df)
    return jsonify({"ok": True})

@app.route("/api/add_barang", methods=["POST"])
def api_add_barang():
    data = request.json or request.form
    kode = (data.get("kode_barang") or "").strip()
    nama = (data.get("nama_barang") or "").strip()
    if not kode or not nama:
        return jsonify({"ok": False, "message": "Kode & nama wajib diisi."}), 400
    stok_df = load_stok()
    if kode in stok_df["Kode Barang"].values:
        return jsonify({"ok": False, "message": "Kode sudah ada."}), 400
    stok_df = pd.concat([stok_df, pd.DataFrame([{"Kode Barang": kode, "Nama Barang": nama, "Saldo": 0}])], ignore_index=True)
    save_stok(stok_df)
    return jsonify({"ok": True})

@app.route("/api/upload_stok_awal", methods=["POST"])
def api_upload_stok_awal():
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "message": "Tidak ada file."}), 400
    df = pd.read_excel(file)
    if not {'Kode Barang', 'Nama Barang', 'Saldo Awal'}.issubset(df.columns):
        return jsonify({"ok": False, "message": "Kolom harus: Kode Barang, Nama Barang, Saldo Awal"}), 400
    stok_df = load_stok()
    mutasi_df = load_mutasi()
    for _, row in df.iterrows():
        kode = str(row["Kode Barang"]).strip()
        nama = str(row["Nama Barang"]).strip()
        saldo_awal = float(row["Saldo Awal"])
        if kode in stok_df["Kode Barang"].values:
            stok_df.loc[stok_df["Kode Barang"] == kode, "Saldo"] = saldo_awal
        else:
            stok_df = pd.concat([stok_df, pd.DataFrame([{"Kode Barang": kode, "Nama Barang": nama, "Saldo": saldo_awal}])], ignore_index=True)
        mutasi_df = pd.concat([mutasi_df, pd.DataFrame([{"Tanggal": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                         "Kode Barang": kode, "Nama Barang": nama,
                                                         "Jenis": "Stok Awal", "Jumlah": saldo_awal,
                                                         "Keterangan": "Upload Stok Awal", "Nama Input": "System"}])], ignore_index=True)
    save_stok(stok_df)
    save_mutasi(mutasi_df)
    return jsonify({"ok": True})

@app.route("/api/export_stok_excel")
def api_export_stok_excel():
    df = load_stok()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Stok")
    output.seek(0)
    return send_file(output, download_name="stok_terkini.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export_mutasi_all")
def api_export_mutasi_all():
    df = load_mutasi()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Mutasi")
    output.seek(0)
    return send_file(output, download_name="riwayat_mutasi.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export_mutasi_period", methods=["POST"])
def api_export_mutasi_period():
    start = request.form.get("start_date")
    end = request.form.get("end_date")
    df = load_mutasi()
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df_filtered = df[(df["Tanggal"] >= start) & (df["Tanggal"] <= end)]
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name="Laporan")
    output.seek(0)
    return send_file(output, download_name=f"laporan_{start}_sd_{end}.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
