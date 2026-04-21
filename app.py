from pathlib import Path
from flask import Flask, jsonify, render_template, request, session, send_file

from reconciliation import ReconciliationService, ReconError

BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__)
app.secret_key = "amco-v12-login-secret"
service = ReconciliationService(BASE_DIR)

APP_USERNAME = "finance"
APP_PASSWORD = "amcosg"


def is_authenticated() -> bool:
    return session.get("authenticated") is True


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/session")
def api_session():
    return jsonify({"authenticated": is_authenticated()})


@app.post("/api/login")
def api_login():
    payload = request.get_json(silent=True) or {}
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", ""))

    if username == APP_USERNAME and password == APP_PASSWORD:
        session["authenticated"] = True
        return jsonify({"ok": True, "message": "Login successful."})

    session.pop("authenticated", None)
    return jsonify({"ok": False, "error": "Invalid username or password."}), 401


@app.post("/api/logout")
def api_logout():
    session.clear()
    return jsonify({"ok": True})



@app.post('/api/export')
def api_export():
    if not is_authenticated():
        return jsonify({'error': 'Please log in first.'}), 401

    sap_file = request.files.get('sap_file')
    os_file = request.files.get('os_file')
    entity = request.form.get('entity', '')
    if not sap_file or not os_file:
        return jsonify({'error': 'Please upload both SAP BFC raw data and OneStream raw data files.'}), 400
    try:
        output = service.export_excel(sap_file, os_file, entity=entity)
        return send_file(
            output,
            as_attachment=True,
            download_name='reconciliation_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except ReconError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': f'Unexpected error: {exc}'}), 500

@app.post("/api/reconcile")
def api_reconcile():
    if not is_authenticated():
        return jsonify({"error": "Please log in first."}), 401

    sap_file = request.files.get("sap_file")
    os_file = request.files.get("os_file")
    if not sap_file or not os_file:
        return jsonify({"error": "Please upload both SAP BFC raw data and OneStream raw data files."}), 400
    try:
        result = service.reconcile(sap_file, os_file)
        return jsonify(result.payload)
    except ReconError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Unexpected error: {exc}"}), 500


if __name__ == "__main__":
    app.run(debug=True)
