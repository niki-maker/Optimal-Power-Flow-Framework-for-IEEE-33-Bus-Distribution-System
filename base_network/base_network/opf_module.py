import opendssdirect as dss
import logging
import pandas as pd
import numpy as np


# ================= METRICS =================
def get_total_real_losses_kw():
    losses = dss.Circuit.Losses()  # [W, var]
    return losses[0] / 1000.0


def get_slack_power_kw():
    dss.Circuit.SetActiveElement("Vsource.Source")
    return abs(dss.CktElement.Powers()[0])  # kW


def get_voltage_deviation():
    vdev = 0.0
    for bus in dss.Circuit.AllBusNames():
        dss.Circuit.SetActiveBus(bus)
        v = dss.Bus.puVmagAngle()[0::2]
        for vi in v:
            vdev += abs(vi - 1.0)
    return vdev


def objective_function(weights=(1.0, 1.0, 1.0)):
    w1, w2, w3 = weights
    loss = get_total_real_losses_kw()
    slack = get_slack_power_kw()
    vdev = get_voltage_deviation()
    return w1 * loss + w2 * slack + w3 * vdev, loss, slack, vdev


# ================= OPF CORE =================
def _run_opf_search(dg_limits, iterations=15, step_kw=50):
    best_J = float("inf")
    best_dispatch = {}

    # Initialize dispatch
    for dg in dg_limits:
        dss.Generators.Name(dg)
        best_dispatch[dg] = dss.Generators.kW()

    for _ in range(iterations):
        for dg, (pmin, pmax) in dg_limits.items():
            dss.Generators.Name(dg)
            current_p = dss.Generators.kW()

            for delta in (-step_kw, step_kw):
                trial_p = max(pmin, min(pmax, current_p + delta))
                dss.Generators.kW(trial_p)

                dss.Text.Command("Solve")
                if not dss.Solution.Converged():
                    continue

                J, *_ = objective_function()
                if J < best_J:
                    best_J = J
                    best_dispatch[dg] = trial_p

            # restore best
            dss.Generators.kW(best_dispatch[dg])

    return best_J, best_dispatch


# ==========================================================
# ✅ SINGLE PUBLIC FUNCTION
# ==========================================================
def run_and_log_opf(
    dg_limits,
    weights=(1.0, 1.0, 0.5),
    iterations=15,
    step_kw=50
):
    """
    Runs OPF end-to-end and logs all results.
    """

    logging.info("🚀 Starting Optimal Power Flow (OPF)...")

    # ---- Base case ----
    base_J, base_loss, base_slack, base_vdev = objective_function(weights)

    # ---- Optimization ----
    best_J, best_dispatch = _run_opf_search(
        dg_limits,
        iterations=iterations,
        step_kw=step_kw
    )

    # ---- Apply optimal dispatch ----
    for dg, p in best_dispatch.items():
        dss.Generators.Name(dg)
        dss.Generators.kW(p)

    dss.Text.Command("Solve")

    # ---- Optimized metrics ----
    opt_J, opt_loss, opt_slack, opt_vdev = objective_function(weights)

    # ---- Logging ----
    logging.info("✅ OPF RESULTS")
    logging.info(f"🔹 Base Loss (kW): {base_loss:.2f}")
    logging.info(f"🔹 OPF  Loss (kW): {opt_loss:.2f}")

    logging.info(f"🔹 Base Slack Import (kW): {base_slack:.2f}")
    logging.info(f"🔹 OPF  Slack Import (kW): {opt_slack:.2f}")

    logging.info(f"🔹 Base Voltage Deviation: {base_vdev:.4f}")
    logging.info(f"🔹 OPF  Voltage Deviation: {opt_vdev:.4f}")

    logging.info(f"⚡ Optimal DG Dispatch: {best_dispatch}")

    return {
        "base": {
            "J": base_J,
            "loss": base_loss,
            "slack": base_slack,
            "vdev": base_vdev,
        },
        "optimal": {
            "J": opt_J,
            "loss": opt_loss,
            "slack": opt_slack,
            "vdev": opt_vdev,
            "dispatch": best_dispatch,
        },
    }
