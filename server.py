from flask import Flask, request, jsonify, render_template_string
import os
from pathlib import Path

from renderer import get_fields, render

app = Flask(__name__, static_folder="static", static_url_path="/static")

_index_html = None


def _get_index_html():
    path = Path(__file__).parent / "templates" / "index.html"
    return path.read_text(encoding="utf-8")


@app.route("/")
def index():
    return render_template_string(_get_index_html())


@app.route("/api/fields")
def api_fields():
    return jsonify(get_fields())


@app.route("/api/preview", methods=["POST"])
def api_preview():
    data = request.get_json()
    if not data or "fields" not in data:
        return jsonify({"error": "fields required"}), 400

    lang = data.get("lang", "ENG").upper()
    if lang not in ("ENG", "BM"):
        return jsonify({"error": "lang must be ENG or BM"}), 400

    try:
        html = render(data["fields"], lang=lang)
        return jsonify({"html": html})
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/render", methods=["POST"])
def api_render():
    data = request.get_json()
    if not data:
        return jsonify({"error": "request body required"}), 400

    if "fields" not in data:
        return jsonify({"error": "'fields' is required"}), 400

    lang = data.get("lang", "ENG").upper()
    if lang not in ("ENG", "BM"):
        return jsonify({"error": "lang must be ENG or BM"}), 400

    try:
        html = render(data["fields"], lang=lang)
        return jsonify({"html": html})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        print(f"\n  NextGen VIP Email Blast — Render Server")
        print(f"  Open http://localhost:5000 in your browser\n")
    app.run(debug=False, port=5000, use_reloader=False)
