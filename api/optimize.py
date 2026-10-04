"""
Vercel entrypoint. The error you hit happens because Vercel's Python
runtime needs a module-level variable literally named `app`, `application`,
or `handler` - a plain script with only `if __name__ == "__main__":` has
nothing for Vercel to call over HTTP.

This file must live at api/optimize.py (Vercel's convention: anything
under /api becomes a serverless function, mapped to /api/<filename>).
It exposes a Flask app - Flask apps are WSGI callables, which is exactly
what Vercel's "Found main.py but it does not export..." message is asking
for.

Route: GET https://<your-project>.vercel.app/api/optimize
Optional query params override assumptions without touching the code, e.g.
  /api/optimize?max_warehouses=4&own_warehouses=true
"""

import os
import sys
from flask import Flask, jsonify, request

# so "from network_optimizer import ..." works regardless of Vercel's cwd
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from network_optimizer import load_data, solve  # noqa: E402

app = Flask(__name__)  # <-- this is the part that was missing

SOURCE_WORKBOOK = os.path.join(
    os.path.dirname(__file__), "..", "Network_Redesign_Optimizer_Model.xlsx"
)


@app.route("/api/optimize", methods=["GET"])
def optimize():
    try:
        data = load_data(SOURCE_WORKBOOK)

        # optional overrides via query string, e.g. ?max_warehouses=4
        if "max_warehouses" in request.args:
            data["assumptions"]["max_warehouses"] = int(request.args["max_warehouses"])
        own_warehouses = request.args.get("own_warehouses", "false").lower() == "true"

        result = solve(data, own_warehouses=own_warehouses)

        wh_totals = {}
        for (w, d), v in result["flows_wd"].items():
            wh_totals[w] = wh_totals.get(w, 0) + v

        top_flows = sorted(result["flows_wd"].items(), key=lambda kv: -kv[1])[:15]

        return jsonify({
            "status": result["status"],
            "total_cost": result["total_cost"],
            "cost_breakdown": result["cost_breakdown"],
            "open_warehouses": result["open_warehouses"],
            "wh_totals": {w: round(v) for w, v in wh_totals.items()},
            "top_flows": [{"warehouse": w, "destination": d, "units": round(v)}
                          for (w, d), v in top_flows],
            "n_destinations": len(data["destinations"]),
        })
    except Exception as e:
        # Vercel swallows stack traces from serverless functions by default,
        # so return the message in the response body to actually see it.
        return jsonify({"error": str(e)}), 500


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True})
