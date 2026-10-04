"""
Network Redesign Optimizer - Python / PuLP version
=====================================================
Replaces the Excel Solver model. Solves the FULL problem (all 208 real
destinations) instead of the 16-point workaround Excel needed to fit under
its 200-decision-variable Solver limit.

Model: two-echelon MILP (Plant -> Warehouse -> Destination)
  - Yi        : open/close each candidate warehouse (binary)
  - W[k,i]    : units shipped from plant k to warehouse i (continuous, >=0)
  - X[i,j]    : units shipped from warehouse i to destination j (continuous, >=0)

Objective: minimize outbound transport + handling + inbound transport + fixed costs

Requires: pip install pulp openpyxl
"""

import openpyxl
import pulp

SOURCE_WORKBOOK = "Network_Redesign_Optimizer_Model.xlsx"  # the file actually committed to this repo


# ---------------------------------------------------------------------------
# 1. LOAD DATA DIRECTLY FROM THE EXCEL WORKBOOK'S DATA SHEETS
#    (these sheets already contain everything; no re-entry needed)
# ---------------------------------------------------------------------------
def load_data(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # --- Destinations: ALL of them, not just the top-16 Excel was limited to
    dest_ws = wb["Demand_by_Destination"]
    destinations, demand = [], {}
    for row in dest_ws.iter_rows(min_row=5, values_only=True):  # header at row 4
        key, name, lat, lon, qty, pct = row
        if key is None or qty in (None, 0) or key == "TOTAL":
            continue
        destinations.append(key)
        demand[key] = qty

    # --- Plants
    plant_ws = wb["Optimizer_Plants"]
    plants, plant_capacity = [], {}
    for row in plant_ws.iter_rows(min_row=5, values_only=True):
        if row[1] is None:
            continue
        _, name, lat, lon, cap = row
        plants.append(name)
        plant_capacity[name] = cap

    # --- Candidate warehouses (names come from the distance matrix header)
    dm_ws = wb["Distance_Matrix"]
    header = [c.value for c in dm_ws[5]]  # row 5 = header row
    warehouses = [h.replace("Candidate: ", "") for h in header if h and h.startswith("Candidate:")]

    # --- Warehouse -> Destination distances (outbound leg), for ALL 208 destinations
    dist_wd = {}  # (warehouse, destination) -> km
    col_start = header.index(f"Candidate: {warehouses[0]}")
    for row in dm_ws.iter_rows(min_row=6, values_only=True):
        key = row[0]
        if key is None or key not in demand:
            continue
        for i, wh in enumerate(warehouses):
            dist_wd[(wh, key)] = row[col_start + i]

    # --- Plant -> Warehouse distances (inbound leg)
    pw_ws = wb["Optimizer_Distance_PW"]
    pw_header = [c.value for c in pw_ws[4]]  # row 4 = header (Plant, then warehouse names)
    dist_pw = {}  # (plant, warehouse) -> km
    for row in pw_ws.iter_rows(min_row=7, values_only=True):
        plant = row[0]
        if plant is None:
            continue
        for i, wh in enumerate(warehouses):
            dist_pw[(plant, wh)] = row[1 + i]

    # --- Assumptions
    in_ws = wb["Optimizer_Inputs"]
    vals = {row[0]: row[1] for row in in_ws.iter_rows(min_row=5, max_row=11, values_only=True)}
    assumptions = {
        "outbound_rate": vals["Outbound freight rate (currency per unit per km) - Warehouse to Destination"],
        "inbound_rate": vals["Inbound freight rate (currency per unit per km) - Plant to Warehouse"],
        "handling_cost": vals["Handling cost (currency per unit, at warehouse)"],
        "fixed_cost": vals["Fixed cost per warehouse, if opened (currency per period)"],
        "wh_capacity": vals["Warehouse capacity (units per period)"],
        "max_warehouses": vals["Maximum number of warehouses to open"],
    }

    return {
        "destinations": destinations, "demand": demand,
        "plants": plants, "plant_capacity": plant_capacity,
        "warehouses": warehouses,
        "dist_wd": dist_wd, "dist_pw": dist_pw,
        "assumptions": assumptions,
    }


# ---------------------------------------------------------------------------
# 2. BUILD AND SOLVE THE MILP
# ---------------------------------------------------------------------------
def solve(data, solver_name="CBC", time_limit=120, own_warehouses=False):
    """
    own_warehouses=False (default): warehouses are third-party / leased space,
        not owned. There is no fixed cost for picking a site - cost scales
        only with volume moved through it (handling fee + transport), which
        is what you actually pay a 3PL. Capacity and handling cost are
        applied as the SAME flat rate to every candidate warehouse
        (a["wh_capacity"], a["handling_cost"]) rather than a per-site number,
        since that's the assumption given for this network.
    own_warehouses=True: adds back the fixed cost term (rent/staff/capex)
        for whichever sites get opened - use this only if some sites are
        company-owned.
    """
    W_set = data["warehouses"]
    D_set = data["destinations"]
    P_set = data["plants"]
    a = data["assumptions"]

    prob = pulp.LpProblem("Network_Redesign", pulp.LpMinimize)

    # Decision variables. Named by numeric index, not by location text --
    # several destination names collapse to the same identifier once punctuation
    # is stripped (e.g. "OBIO AKPOR" / "OBIO-AKPOR" / "OBIO/AKPOR"), which produced
    # silent variable name collisions and a corrupted LP file when names were used.
    w_idx = {w: i for i, w in enumerate(W_set)}
    d_idx = {d: i for i, d in enumerate(D_set)}
    p_idx = {p: i for i, p in enumerate(P_set)}

    Y = {w: pulp.LpVariable(f"Open_{w_idx[w]}", cat="Binary") for w in W_set}
    X = {(w, d): pulp.LpVariable(f"Flow_WD_{w_idx[w]}_{d_idx[d]}", lowBound=0) for w in W_set for d in D_set}
    Wf = {(p, w): pulp.LpVariable(f"Flow_PW_{p_idx[p]}_{w_idx[w]}", lowBound=0) for p in P_set for w in W_set}

    # Objective. handling_cost and wh_capacity are single flat numbers applied
    # to every warehouse (a["handling_cost"], a["wh_capacity"]) - not a
    # per-site lookup - because all candidate sites are assumed identical on
    # both counts.
    outbound_cost = pulp.lpSum(X[(w, d)] * data["dist_wd"][(w, d)] * a["outbound_rate"] for w in W_set for d in D_set)
    handling_cost = pulp.lpSum(X[(w, d)] * a["handling_cost"] for w in W_set for d in D_set)
    inbound_cost = pulp.lpSum(Wf[(p, w)] * data["dist_pw"][(p, w)] * a["inbound_rate"] for p in P_set for w in W_set)

    if own_warehouses:
        fixed_cost = pulp.lpSum(Y[w] * a["fixed_cost"] for w in W_set)
        prob += outbound_cost + handling_cost + inbound_cost + fixed_cost
    else:
        # No fixed/capex cost: these are third-party warehouses, paid for
        # only by what passes through them (handling_cost already covers
        # that). Opening one "costs" nothing extra beyond its throughput.
        fixed_cost = pulp.lpSum(0 * Y[w] for w in W_set)  # kept as an LpAffineExpression, always 0
        prob += outbound_cost + handling_cost + inbound_cost

    # Constraints
    # Note: constraint names are indexed rather than built from destination text,
    # since several destination names (e.g. "OBIO AKPOR" / "OBIO-AKPOR" / "OBIO/AKPOR")
    # collapse to the same sanitized identifier and PuLP requires unique names.
    for idx, d in enumerate(D_set):  # 1. demand satisfied exactly
        prob += pulp.lpSum(X[(w, d)] for w in W_set) == data["demand"][d], f"Demand_{idx}"

    for idx, w in enumerate(W_set):  # 2. outbound capacity + 3. throughput balance
        prob += pulp.lpSum(X[(w, d)] for d in D_set) <= a["wh_capacity"] * Y[w], f"OutCap_{idx}"
        prob += pulp.lpSum(Wf[(p, w)] for p in P_set) == pulp.lpSum(X[(w, d)] for d in D_set), f"Balance_{idx}"

    for idx, p in enumerate(P_set):  # 4. plant capacity
        prob += pulp.lpSum(Wf[(p, w)] for w in W_set) <= data["plant_capacity"][p], f"PlantCap_{idx}"

    prob += pulp.lpSum(Y[w] for w in W_set) <= a["max_warehouses"], "MaxWarehouses"  # 8. max open
    # (5. binary and 6/7. non-negativity are enforced by variable definitions above,
    #  not bolt-on constraints -- this is the part Excel couldn't guarantee reliably)

    solver = pulp.PULP_CBC_CMD(msg=1, timeLimit=time_limit)
    prob.solve(solver)

    return {
        "status": pulp.LpStatus[prob.status],
        "total_cost": pulp.value(prob.objective),
        "cost_breakdown": {
            "outbound_transport": pulp.value(outbound_cost),
            "handling": pulp.value(handling_cost),
            "inbound_transport": pulp.value(inbound_cost),
            "fixed": pulp.value(fixed_cost),
        },
        "open_warehouses": [w for w in W_set if Y[w].value() > 0.5],
        "flows_wd": {(w, d): X[(w, d)].value() for w in W_set for d in D_set if X[(w, d)].value() > 1e-6},
        "flows_pw": {(p, w): Wf[(p, w)].value() for p in P_set for w in W_set if Wf[(p, w)].value() > 1e-6},
    }


# ---------------------------------------------------------------------------
# 3. RUN
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    data = load_data(SOURCE_WORKBOOK)
    print(f"Loaded {len(data['destinations'])} destinations (vs. 16 in the Excel version), "
          f"{len(data['plants'])} plants, {len(data['warehouses'])} candidate warehouses.")
    n_vars = len(data["warehouses"]) + len(data["plants"]) * len(data["warehouses"]) + \
             len(data["warehouses"]) * len(data["destinations"])
    print(f"Decision variables: {n_vars:,}  (Excel's free Solver caps out at 200)")

    # own_warehouses=False: no fixed/capex cost, since these are third-party
    # warehouses paid for by throughput only. Capacity and handling cost are
    # the same flat number for every candidate site (see Optimizer_Inputs).
    result = solve(data, own_warehouses=False)

    print("\n--- RESULT (third-party warehouses, no fixed cost) ---")
    print("Status:", result["status"])
    print(f"Total network cost: {result['total_cost']:,.0f}")
    for k, v in result["cost_breakdown"].items():
        print(f"  {k}: {v:,.0f}")
    print("Open warehouses:", result["open_warehouses"])
