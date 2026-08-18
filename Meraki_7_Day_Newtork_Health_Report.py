#!/usr/bin/env python3
"""
Meraki Organization Monthly Health Report
==========================================

Pulls 30-day telemetry across every product type in a Meraki organization
(appliance/MX, switch/MS, wireless/MR, camera/MV, sensor/MT, cellularGateway/Z)
and produces:

  1. A full raw-data JSON export (for archiving / further analysis)
  2. An executive-ready HTML summary report
  3. A flat CSV summary (one row per finding, easy to drop into Excel/Sheets)

Requirements:
  pip install meraki --break-system-packages

Usage:
  set MERAKI_DASHBOARD_API_KEY=your_key_here      (Windows cmd)
  python meraki_monthly_health_report.py

  or pass an org id directly to skip the picker:
  python meraki_monthly_health_report.py --org-id 123456
"""

import os
import sys
import json
import csv
import time
import argparse
from datetime import datetime, timezone, timedelta
from collections import defaultdict

try:
    import meraki
    from meraki.exceptions import APIError
except ImportError:
    print("ERROR: the 'meraki' package is not installed.")
    print("Run:  pip install meraki --break-system-packages")
    sys.exit(1)

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# CONFIG - EDIT THESE TWO LINES FOR EACH ORG YOU RUN THIS AGAINST
# ---------------------------------------------------------------------------
API_KEY = os.environ.get("MERAKI_DASHBOARD_API_KEY", "PASTE_YOUR_MERAKI_API_KEY_HERE")
ORG_ID = "361642"   # e.g. "361642"  -- leave as "" to get an interactive picker instead
LOCAL_UTC_OFFSET_HOURS = -4          # e.g. -4 for EDT, -5 for EST/CDT, +8 for PHT - adjust for the site's timezone
# ---------------------------------------------------------------------------

TIMESPAN_SECONDS = 604800  # 7 days
OUTPUT_DIR = "."
TIMESTAMP_TAG = datetime.now().strftime("%Y-%m-%d_%H%M")
JSON_OUT = os.path.join(OUTPUT_DIR, f"meraki_health_report_{TIMESTAMP_TAG}.json")
HTML_OUT = os.path.join(OUTPUT_DIR, f"meraki_health_report_{TIMESTAMP_TAG}.html")
CSV_OUT = os.path.join(OUTPUT_DIR, f"meraki_health_report_{TIMESTAMP_TAG}.csv")
# Stable, un-timestamped copy + a tiny manifest file. These are what make the
# "refresh" button in the HTML meaningful: teammates always load the SAME
# URL (HTML_LATEST), and the page polls MANIFEST_OUT to know whether a newer
# run has landed in the repo since they opened it. See write_html() below.
HTML_LATEST = os.path.join(OUTPUT_DIR, "meraki_health_report_latest.html")
MANIFEST_OUT = os.path.join(OUTPUT_DIR, "latest.json")
LOCAL_TZ = timezone(timedelta(hours=LOCAL_UTC_OFFSET_HOURS))

# High-error thresholds used for flagging (tune as needed)
CRC_ERROR_THRESHOLD = 100
POE_SATURATION_PCT = 85
WIRELESS_RETRY_PCT_THRESHOLD = 20
LOW_BATTERY_PCT = 20
APPLIANCE_UTILIZATION_THRESHOLD_PCT = 70   # flag MX as "near capacity" above this
SWITCH_PORT_UTILIZATION_THRESHOLD_PCT = 80 # flag MS ports as "near capacity" above this (traffic / rated link speed)
AP_CHANNEL_UTILIZATION_THRESHOLD_PCT = 60  # flag MR radios as "congested" above this (RF channel utilization %)
LICENSE_EXPIRING_SOON_DAYS = 30            # flag licenses expiring within this window
NON_GIGABIT_SPEEDS = ("10 Mbps", "100 Mbps", "10Mbps", "100Mbps")  # crude "degraded uplink speed" match


def _speed_to_kbps(speed_str):
    """Converts Meraki's port speed string (e.g. '1 Gbps', '100 Mbps') to kbps
    so it can be compared against measured traffic. Returns None for anything
    that can't be parsed (port down, no link, unexpected format, etc.) so
    callers can skip utilization math rather than divide by a wrong number."""
    if not speed_str or not isinstance(speed_str, str):
        return None
    s = speed_str.strip().lower()
    try:
        if "gbps" in s:
            return float(s.replace("gbps", "").strip()) * 1_000_000
        if "mbps" in s:
            return float(s.replace("mbps", "").strip()) * 1_000
    except ValueError:
        return None
    return None


def format_local_str(dt_utc):
    """Formats a UTC datetime as mm-dd-yyyy 12hr AM/PM in the configured local timezone.
    Uses a single explicit fixed-offset conversion (astimezone) so this does NOT
    depend on, and cannot be double-shifted by, the machine's own system timezone."""
    if not dt_utc:
        return ""
    return dt_utc.astimezone(LOCAL_TZ).strftime("%m-%d-%Y %I:%M %p")


# ---------------------------------------------------------------------------
# CONSOLE HELPERS
# ---------------------------------------------------------------------------
def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")


def progress(current, total, label=""):
    pct = int((current / total) * 100) if total else 100
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    sys.stdout.write(f"\r  [{bar}] {pct:3d}%  {label}")
    sys.stdout.flush()
    if current >= total:
        sys.stdout.write("\n")


# ---------------------------------------------------------------------------
# RETRY / RATE-LIMIT WRAPPER
# ---------------------------------------------------------------------------
def call_with_retry(func, *args, max_attempts=5, base_delay=2, **kwargs):
    """
    Calls a Meraki SDK function with exponential backoff on 429s and
    transient 5xx errors. The SDK itself already retries 429s internally
    (wait_on_rate_limit=True below), but this wrapper adds a second layer
    of resilience for anything that slips through, plus graceful handling
    for endpoints that don't apply to a given org/network (404/400).
    """
    attempt = 0
    while True:
        try:
            return func(*args, **kwargs)
        except APIError as e:
            status = getattr(e.response, "status_code", None)
            attempt += 1
            if status == 429 and attempt <= max_attempts:
                delay = base_delay * (2 ** (attempt - 1))
                log(f"  Rate limited (429). Backing off {delay}s (attempt {attempt}/{max_attempts})...")
                time.sleep(delay)
                continue
            if status in (400, 404):
                # Endpoint not applicable (e.g. no cameras in network) - skip quietly
                return None
            if status and status >= 500 and attempt <= max_attempts:
                delay = base_delay * (2 ** (attempt - 1))
                log(f"  Server error ({status}). Retrying in {delay}s...")
                time.sleep(delay)
                continue
            log(f"  API error on {func.__name__ if hasattr(func, '__name__') else func}: {e}")
            return None
        except Exception as e:
            log(f"  Unexpected error: {e}")
            return None


# ---------------------------------------------------------------------------
# ORG SELECTION
# ---------------------------------------------------------------------------
def choose_organization(dashboard, org_id_arg=None):
    orgs = call_with_retry(dashboard.organizations.getOrganizations) or []
    if not orgs:
        log("No organizations found for this API key. Exiting.")
        sys.exit(1)

    if org_id_arg:
        match = next((o for o in orgs if str(o["id"]) == str(org_id_arg)), None)
        if not match:
            log(f"Org ID {org_id_arg} not found for this API key.")
            sys.exit(1)
        return match

    if len(orgs) == 1:
        log(f"Only one organization available: {orgs[0]['name']}")
        return orgs[0]

    print("\nAvailable organizations:")
    for i, o in enumerate(orgs, 1):
        print(f"  {i}. {o['name']}  (ID: {o['id']})")
    while True:
        choice = input(f"\nSelect organization [1-{len(orgs)}]: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(orgs):
            return orgs[int(choice) - 1]
        print("Invalid choice, try again.")


# ---------------------------------------------------------------------------
# 1. DEVICE INVENTORY
# ---------------------------------------------------------------------------
def get_all_devices(dashboard, org_id):
    log("Fetching device inventory...")
    devices = call_with_retry(
        dashboard.organizations.getOrganizationDevices, org_id, total_pages="all"
    ) or []
    by_type = defaultdict(list)
    for d in devices:
        by_type[d.get("productType", "unknown")].append(d)
    log(f"  Found {len(devices)} devices: " +
        ", ".join(f"{k}={len(v)}" for k, v in by_type.items()))
    return devices, by_type


# ---------------------------------------------------------------------------
# 2. AVAILABILITY / OUTAGES (all product types)
# ---------------------------------------------------------------------------
def extract_status_value(detail_list):
    if not isinstance(detail_list, list):
        return None
    for item in detail_list:
        if item.get("name") == "status":
            return item.get("value")
    return None


def get_availability_data(dashboard, org_id):
    log("Fetching current device availability (all product types)...")
    current = call_with_retry(
        dashboard.organizations.getOrganizationDevicesAvailabilities,
        org_id, total_pages="all"
    ) or []

    log("Fetching 30-day availability change history (all product types)...")
    history = call_with_retry(
        dashboard.organizations.getOrganizationDevicesAvailabilitiesChangeHistory,
        org_id, total_pages="all", timespan=TIMESPAN_SECONDS
    ) or []

    log(f"  Current snapshot: {len(current)} devices | History: {len(history)} events")

    # Compute per-device downtime totals
    events_by_serial = defaultdict(list)
    for ev in history:
        serial = ev.get("device", {}).get("serial")
        ts = ev.get("ts")
        if not serial or not ts:
            continue
        dt = _parse_ts(ts)
        old_status = extract_status_value(ev.get("details", {}).get("old", []))
        new_status = extract_status_value(ev.get("details", {}).get("new", []))
        events_by_serial[serial].append({"dt": dt, "old": old_status, "new": new_status, "raw": ev})

    downtime_summary = []
    last_offline_transition_by_serial = {}
    network_name_by_serial = {}
    for serial, events in events_by_serial.items():
        events.sort(key=lambda e: e["dt"] or datetime.min.replace(tzinfo=timezone.utc))
        device_name = events[0]["raw"].get("device", {}).get("name", serial)
        product_type = events[0]["raw"].get("device", {}).get("productType", "unknown")
        network_name = events[0]["raw"].get("network", {}).get("name", "")
        network_name_by_serial[serial] = network_name
        offline_start = None
        total_downtime = 0
        outage_count = 0
        for ev in events:
            if ev["new"] == "offline":
                offline_start = ev["dt"]
                outage_count += 1
                last_offline_transition_by_serial[serial] = ev["dt"]
            elif ev["new"] == "online" and offline_start:
                if ev["dt"]:
                    total_downtime += (ev["dt"] - offline_start).total_seconds()
                offline_start = None
        if offline_start:
            total_downtime += (datetime.now(timezone.utc) - offline_start).total_seconds()

        downtime_summary.append({
            "serial": serial,
            "name": device_name,
            "productType": product_type,
            "network": network_name,
            "outageCount": outage_count,
            "totalDowntimeSeconds": round(total_downtime),
            "totalDowntimeHours": round(total_downtime / 3600, 2),
        })

    return {
        "current_snapshot": current,
        "change_history": history,
        "downtime_summary": sorted(downtime_summary, key=lambda x: -x["totalDowntimeSeconds"]),
        "last_offline_transition_by_serial": last_offline_transition_by_serial,
        "network_name_by_serial": network_name_by_serial,
    }


def _parse_ts(iso_str):
    if not iso_str:
        return None
    try:
        clean = iso_str.split(".")[0].replace("Z", "")
        return datetime.strptime(clean, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 3. WAN / APPLIANCE HEALTH
# ---------------------------------------------------------------------------
def get_wan_health(dashboard, org_id):
    log("Fetching WAN loss/latency/jitter (uplinks)...")
    loss_latency = call_with_retry(
        dashboard.organizations.getOrganizationDevicesUplinksLossAndLatency,
        org_id, timespan=TIMESPAN_SECONDS
    ) or []

    log("Fetching WAN uplink statuses (current)...")
    uplink_statuses = call_with_retry(
        dashboard.appliance.getOrganizationApplianceUplinkStatuses,
        org_id, total_pages="all"
    ) or []

    log("Fetching WAN bandwidth usage history...")
    usage_history = call_with_retry(
        dashboard.appliance.getOrganizationApplianceUplinksUsageByNetwork,
        org_id, timespan=TIMESPAN_SECONDS
    ) or []

    # Flag high loss/latency uplinks
    flagged = []
    for entry in loss_latency:
        for series_point in entry.get("timeSeries", []):
            loss_pct = series_point.get("lossPercent")
            latency_ms = series_point.get("latencyMs")
            if (loss_pct and loss_pct > 2) or (latency_ms and latency_ms > 150):
                flagged.append({
                    "serial": entry.get("serial"),
                    "network": entry.get("network", {}).get("name"),
                    "uplink": entry.get("uplink"),
                    "ts": series_point.get("ts"),
                    "lossPercent": loss_pct,
                    "latencyMs": latency_ms,
                })
                break  # one flag per uplink is enough for the summary

    return {
        "loss_latency_raw": loss_latency,
        "uplink_statuses": uplink_statuses,
        "usage_history": usage_history,
        "flagged_uplinks": flagged,
    }


# ---------------------------------------------------------------------------
# 4. SWITCH HEALTH
# ---------------------------------------------------------------------------
def get_switch_health(dashboard, org_id, switch_devices):
    if not switch_devices:
        log("No switches in this organization - skipping switch health.")
        return {"port_statuses": [], "flagged_ports": [], "flagged_high_utilization_ports": []}

    log(f"Fetching switch port statuses for {len(switch_devices)} switch(es)...")
    raw_response = call_with_retry(
        dashboard.switch.getOrganizationSwitchPortsStatusesBySwitch,
        org_id, total_pages="all", perPage=20
    )
    # This endpoint returns {"items": [...], "meta": {...}} rather than a
    # plain list - unwrap it if needed so we always end up with a list.
    if isinstance(raw_response, dict):
        port_statuses = raw_response.get("items", [])
    elif isinstance(raw_response, list):
        port_statuses = raw_response
    else:
        port_statuses = []

    # Errors/warnings (CRC, FCS, framing, collisions, etc.) are physical-layer
    # QUALITY problems, not a capacity/utilization signal - a port can throw
    # errors at 5% load or run at 100% with zero errors. Keep these under
    # switch health, separate from the real utilization numbers below.
    flagged_ports = []
    # Real capacity signal: measured traffic vs. this port's rated link speed.
    flagged_high_utilization_ports = []
    for switch_entry in port_statuses:
        switch_name = switch_entry.get("name", switch_entry.get("serial"))
        network_name = switch_entry.get("network", {}).get("name", "")
        for port in switch_entry.get("ports", []):
            errors = port.get("errors", []) or []
            warnings = port.get("warnings", []) or []
            poe_info = port.get("poe", {}) or {}
            is_uplink = port.get("isUplink", False)
            if errors or warnings:
                flagged_ports.append({
                    "switch": switch_name,
                    "serial": switch_entry.get("serial"),
                    "network": network_name,
                    "portId": port.get("portId"),
                    "isUplink": is_uplink,
                    "status": port.get("status"),
                    "speed": port.get("speed"),
                    "duplex": port.get("duplex"),
                    "errors": errors,
                    "warnings": warnings,
                    "poeEnabled": poe_info.get("isAllocated", False),
                })

            speed_kbps = _speed_to_kbps(port.get("speed"))
            traffic_kbps = (port.get("trafficInKbps") or {}).get("total")
            if speed_kbps and traffic_kbps is not None:
                util_pct = round((traffic_kbps / speed_kbps) * 100, 1)
                if util_pct >= SWITCH_PORT_UTILIZATION_THRESHOLD_PCT:
                    flagged_high_utilization_ports.append({
                        "switch": switch_name,
                        "serial": switch_entry.get("serial"),
                        "network": network_name,
                        "portId": port.get("portId"),
                        "isUplink": is_uplink,
                        "speed": port.get("speed"),
                        "utilizationPercent": util_pct,
                    })

    return {
        "port_statuses": port_statuses,
        "flagged_ports": flagged_ports,
        "flagged_high_utilization_ports": flagged_high_utilization_ports,
    }


# ---------------------------------------------------------------------------
# 5. WIRELESS HEALTH
# ---------------------------------------------------------------------------
def get_wireless_health(dashboard, org_id, networks):
    wireless_networks = [
        n for n in networks
        if "wireless" in n.get("productTypes", [])
    ]
    if not wireless_networks:
        log("No wireless networks - skipping Wi-Fi health.")
        return {"connection_stats_by_network": [], "flagged_networks": []}

    log(f"Fetching wireless connection stats for {len(wireless_networks)} network(s)...")
    results = []
    flagged = []
    for i, net in enumerate(wireless_networks, 1):
        stats = call_with_retry(
            dashboard.wireless.getNetworkWirelessConnectionStats,
            net["id"], timespan=TIMESPAN_SECONDS
        )
        progress(i, len(wireless_networks), net.get("name", ""))
        if not stats:
            continue
        results.append({"network": net.get("name"), "networkId": net["id"], "stats": stats})

        total_attempts = sum(stats.get(k, 0) for k in ("assoc", "auth", "dhcp", "dns", "success"))
        failures = total_attempts - stats.get("success", 0) if total_attempts else 0
        failure_pct = (failures / total_attempts * 100) if total_attempts else 0
        if failure_pct > WIRELESS_RETRY_PCT_THRESHOLD:
            flagged.append({
                "network": net.get("name"),
                "networkId": net["id"],
                "failurePercent": round(failure_pct, 1),
                "stats": stats,
            })

    return {"connection_stats_by_network": results, "flagged_networks": flagged}


# ---------------------------------------------------------------------------
# 5b. WIRELESS RF CHANNEL UTILIZATION (real capacity/congestion signal)
#     Connection-failure rate (above) tells you "something's wrong" but not
#     reliably "this AP is out of capacity" - a bad DHCP scope or flaky
#     RADIUS server produces the same symptom. This endpoint measures how
#     saturated the RF spectrum actually is at each AP/band, which is the
#     direct "is this AP overloaded" signal.
# ---------------------------------------------------------------------------
def get_ap_channel_utilization(dashboard, org_id, wireless_devices, network_id_to_name):
    if not wireless_devices:
        log("No access points in this organization - skipping RF channel utilization.")
        return {"by_device": [], "flagged_aps": []}

    log(f"Fetching RF channel utilization for {len(wireless_devices)} AP(s)...")
    raw = call_with_retry(
        dashboard.wireless.getOrganizationWirelessDevicesChannelUtilizationByDevice,
        org_id, total_pages="all", timespan=TIMESPAN_SECONDS
    ) or []

    name_by_serial = {d.get("serial"): d.get("name") or d.get("serial") for d in wireless_devices}

    by_device = []
    flagged = []
    for entry in raw:
        serial = entry.get("serial")
        net_id = entry.get("network", {}).get("id", "")
        network_name = entry.get("network", {}).get("name") or network_id_to_name.get(net_id, net_id)
        device_name = name_by_serial.get(serial, serial)
        for band_entry in entry.get("byBand", []):
            total_pct = (band_entry.get("total") or {}).get("percentage")
            if total_pct is None:
                continue
            row = {
                "serial": serial,
                "name": device_name,
                "network": network_name,
                "band": band_entry.get("band"),
                "utilizationPercent": total_pct,
            }
            by_device.append(row)
            if total_pct >= AP_CHANNEL_UTILIZATION_THRESHOLD_PCT:
                flagged.append(row)

    return {"by_device": by_device, "flagged_aps": flagged}


# ---------------------------------------------------------------------------
# 6. CAMERA HEALTH  (built from availability data - Meraki has no separate
#    org-wide "camera health score" endpoint, so uptime/outage counts from
#    section 2, filtered to productType == 'camera', are the source of truth)
# ---------------------------------------------------------------------------
def get_camera_health(availability_data):
    camera_downtime = [
        d for d in availability_data["downtime_summary"] if d["productType"] == "camera"
    ]
    camera_current = [
        d for d in availability_data["current_snapshot"] if d.get("productType") == "camera"
    ]
    offline_now = [d for d in camera_current if d.get("status") == "offline"]
    return {
        "total_cameras": len(camera_current),
        "offline_now": offline_now,
        "downtime_summary": camera_downtime,
    }


# ---------------------------------------------------------------------------
# 7. MT SENSOR HEALTH
# ---------------------------------------------------------------------------
def get_sensor_health(dashboard, org_id, sensor_devices):
    if not sensor_devices:
        log("No MT sensors in this organization - skipping sensor health.")
        return {"latest_readings": [], "low_battery": []}

    log(f"Fetching latest readings for {len(sensor_devices)} sensor(s)...")
    readings = call_with_retry(
        dashboard.sensor.getOrganizationSensorReadingsLatest,
        org_id, total_pages="all"
    ) or []

    low_battery = []
    for r in readings:
        for reading in r.get("readings", []):
            if reading.get("metric") == "battery":
                pct = reading.get("battery", {}).get("percentage")
                if pct is not None and pct < LOW_BATTERY_PCT:
                    low_battery.append({
                        "serial": r.get("serial"),
                        "network": r.get("network", {}).get("name"),
                        "batteryPercent": pct,
                    })

    return {"latest_readings": readings, "low_battery": low_battery}


# ---------------------------------------------------------------------------
# 8. FIRMWARE COMPLIANCE
#     IMPORTANT: getOrganizationFirmwareUpgrades (the endpoint this used to
#     call) only returns a log of upgrade EVENTS - each entry's "status" is
#     just "Cancelled" or "Completed", and there's no currentVersion/
#     isUpgradeAvailable data on it at all. That's why this section was
#     always empty: it was filtering for a "nextUpgrade.toVersion" field
#     that endpoint doesn't provide except for stuff that's already been
#     scheduled - which for most orgs is nothing, since most sites showing
#     "Recommended/Warning" in the dashboard haven't had an upgrade
#     scheduled yet (see the "Upgrade scheduled: No" column in the
#     dashboard's Firmware Upgrades > Schedule Upgrades view).
#     The correct data lives on getNetworkFirmwareUpgrades - a per-network
#     call that reports isUpgradeAvailable/availableVersions/currentVersion
#     per product regardless of whether anything's been scheduled - so this
#     now loops every network to build the same picture as that dashboard
#     screen. There's no organization-wide equivalent for this data.
# ---------------------------------------------------------------------------
def get_firmware_compliance(dashboard, org_id, devices, networks):
    log(f"Auditing firmware versions across {len(networks)} network(s)...")

    # Map (networkId, productType) -> device count, so we can turn
    # per-network firmware entries into an actual affected-device count.
    device_count_by_net_product = defaultdict(int)
    for d in devices:
        device_count_by_net_product[(d.get("networkId"), d.get("productType"))] += 1

    outdated = []
    total_devices_with_upgrade_available = 0
    raw_by_network = []

    for i, net in enumerate(networks, 1):
        net_id = net["id"]
        net_name = net.get("name", net_id)
        info = call_with_retry(dashboard.networks.getNetworkFirmwareUpgrades, net_id)
        progress(i, len(networks), net_name)
        if not info:
            continue  # network doesn't support firmware management (e.g. SM-only) - skip quietly
        raw_by_network.append({"network": net_name, "networkId": net_id, "info": info})

        products = info.get("products", {}) or {}
        for product_type, p_info in products.items():
            if not p_info or not p_info.get("isUpgradeAvailable"):
                continue

            current = p_info.get("currentVersion") or {}
            available_versions = p_info.get("availableVersions") or []
            # Meraki's dashboard labels the GA build "Recommended" - that's
            # the stable-release entry. Prefer it; fall back to whatever's
            # offered (e.g. a beta/candidate build) if no stable is listed.
            recommended = next((v for v in available_versions if v.get("releaseType") == "stable"), None)
            recommended = recommended or (available_versions[0] if available_versions else {})
            next_up = p_info.get("nextUpgrade") or {}

            affected_device_count = device_count_by_net_product.get((net_id, product_type), 0)
            total_devices_with_upgrade_available += affected_device_count
            outdated.append({
                "productType": product_type,
                "network": net_name,
                "currentVersion": current.get("shortName"),
                "availableVersion": recommended.get("shortName"),
                "releaseType": recommended.get("releaseType"),
                "scheduledFor": next_up.get("time"),  # None/absent if not yet scheduled
                "affectedDeviceCount": affected_device_count,
            })

    return {
        "raw": raw_by_network,
        "outdated_summary": outdated,
        "total_devices_with_upgrade_available": total_devices_with_upgrade_available,
    }


# ---------------------------------------------------------------------------
# 8b. FIRMWARE COMPLIANCE - ROLLED UP BY NETWORK
#     The per-product table above answers "which product on which network is
#     behind." This answers the more skimmable executive question: "which
#     sites have anything pending at all." Pure rollup of the same data
#     already fetched above - no extra API call.
# ---------------------------------------------------------------------------
def get_networks_with_firmware_upgrades(firmware_data):
    by_network = defaultdict(lambda: {"products": [], "totalDevicesAffected": 0, "earliestScheduled": None})
    for row in firmware_data["outdated_summary"]:
        net_name = row["network"] or "Unknown"
        entry = by_network[net_name]
        entry["products"].append(f"{row['productType']} ({row['currentVersion']} \u2192 {row['availableVersion']})")
        entry["totalDevicesAffected"] += row["affectedDeviceCount"]
        scheduled_for = row.get("scheduledFor")
        if scheduled_for and (not entry["earliestScheduled"] or scheduled_for < entry["earliestScheduled"]):
            entry["earliestScheduled"] = scheduled_for

    rollup = []
    for net_name, data in by_network.items():
        rollup.append({
            "network": net_name,
            "productsSummary": "; ".join(data["products"]),
            "totalDevicesAffected": data["totalDevicesAffected"],
            "scheduledFor": data["earliestScheduled"] or "Not scheduled",
        })
    rollup.sort(key=lambda r: -r["totalDevicesAffected"])
    return rollup


# ---------------------------------------------------------------------------
# 12. OFFLINE DEVICES BROKEN DOWN BY TYPE (MX / MS / MR / Camera / Sensor)
# ---------------------------------------------------------------------------
PRODUCT_TYPE_LABELS = {
    "appliance": "MX (Security Appliance)",
    "switch": "MS (Switch)",
    "wireless": "MR (Access Point)",
    "camera": "MV (Camera)",
    "sensor": "MT (Sensor)",
    "cellularGateway": "MG / Z (Cellular Gateway)",
}


def get_offline_by_type(availability_data, network_id_to_name):
    current = availability_data["current_snapshot"]
    last_offline_transition = availability_data.get("last_offline_transition_by_serial", {})
    network_name_by_serial = availability_data.get("network_name_by_serial", {})
    downtime_by_serial = {d["serial"]: d for d in availability_data["downtime_summary"]}

    offline_by_type = defaultdict(int)
    alerting_by_type = defaultdict(int)
    offline_devices_detail = []

    for d in current:
        pt = d.get("productType", "unknown")
        status = d.get("status")
        if status == "offline":
            offline_by_type[pt] += 1
        elif status == "alerting":
            alerting_by_type[pt] += 1

        if status != "offline":
            continue

        serial = d.get("serial", "")
        # Network name resolution priority: history-derived name -> org network list -> raw id
        net_id = d.get("network", {}).get("id", "")
        network_name = (
            network_name_by_serial.get(serial)
            or network_id_to_name.get(net_id)
            or net_id
            or "Unknown"
        )

        last_offline_dt = last_offline_transition.get(serial)
        downtime_entry = downtime_by_serial.get(serial)

        if last_offline_dt:
            last_seen_display = f"Went offline {format_local_str(last_offline_dt)}"
        else:
            last_seen_display = f"Offline before this {TIMESPAN_SECONDS // 86400}-day window"

        downtime_display = f"{downtime_entry['totalDowntimeHours']}h this window" if downtime_entry else "Unknown"

        offline_devices_detail.append({
            "deviceTypeLabel": PRODUCT_TYPE_LABELS.get(pt, pt),
            "productType": pt,
            "name": d.get("name") or serial,
            "network": network_name,
            "serial": serial,
            "lastSeen": last_seen_display,
            "downtime": downtime_display,
        })

    offline_devices_detail.sort(key=lambda x: x["deviceTypeLabel"])

    return {
        "offline_by_type": {PRODUCT_TYPE_LABELS.get(k, k): v for k, v in offline_by_type.items()},
        "alerting_by_type": {PRODUCT_TYPE_LABELS.get(k, k): v for k, v in alerting_by_type.items()},
        "offline_devices_detail": offline_devices_detail,
    }


# ---------------------------------------------------------------------------
# 13a. MX (APPLIANCE) UTILIZATION - dedicated capacity-sizing view
#     Meraki's getOrganizationSummaryTopAppliancesByUtilization is the only
#     product type with a true, hardware-rated utilization percentage
#     (current throughput vs. that model's rated max throughput). This is
#     kept as its own section - rather than folded into the generic risk
#     list below - because its purpose is different: it's meant to answer
#     "does this site's MX need a bigger model," so it shows every appliance
#     the endpoint returns (not just ones over the alert threshold) so you
#     can see the whole trend, not just who's currently flagged.
#     NOTE: this Meraki endpoint only returns the top 10 appliances by
#     utilization org-wide - for orgs with more than 10 MX's, appliances
#     outside the top 10 (i.e. the least-utilized ones) won't appear here.
# ---------------------------------------------------------------------------
def get_mx_utilization_overview(dashboard, org_id):
    log("Fetching MX appliance utilization (capacity/upgrade sizing)...")
    top_appliances = call_with_retry(
        dashboard.organizations.getOrganizationSummaryTopAppliancesByUtilization,
        org_id, timespan=TIMESPAN_SECONDS
    ) or []

    rows = []
    for a in top_appliances:
        pct = a.get("utilization", {}).get("average", {}).get("percentage")
        rows.append({
            "name": a.get("name") or a.get("serial"),
            "serial": a.get("serial"),
            "model": a.get("model"),
            "network": a.get("network", {}).get("name"),
            "utilizationPercent": pct,
            "nearCapacity": pct is not None and pct >= APPLIANCE_UTILIZATION_THRESHOLD_PCT,
        })
    rows.sort(key=lambda r: (r["utilizationPercent"] is None, -(r["utilizationPercent"] or 0)))

    flagged_mx = [r for r in rows if r["nearCapacity"]]

    return {"all_appliances": rows, "flagged_near_capacity": flagged_mx}


# ---------------------------------------------------------------------------
# 13b. DEVICE HEALTH & RISK SIGNALS - MS and MR real capacity metrics
#     MX has its own dedicated section above. For MS and MR, this now uses
#     genuine capacity/congestion measurements instead of proxies:
#       MS -> per-port utilization % (measured traffic / rated port speed)
#       MR -> RF channel utilization % (how saturated the spectrum is at
#             that AP/band - Meraki's direct "is this AP overloaded" metric)
#     Camera "offline" and Sensor "low battery" are intentionally NOT
#     included here - they're availability/maintenance issues, not
#     capacity signals, and they're already covered under the Offline
#     Devices and Environmental Sensors sections respectively.
# ---------------------------------------------------------------------------
def get_device_health_risk_signals(switch_data, ap_channel_util_data):
    combined_risk_items = []

    for p in switch_data.get("flagged_high_utilization_ports", []):
        combined_risk_items.append({
            "deviceType": "MS",
            "item": f"{p['switch']} - port {p['portId']}",
            "network": p.get("network", ""),
            "metric": "Port Utilization (traffic / rated speed)",
            "value": f"{p['utilizationPercent']}%",
        })

    for a in ap_channel_util_data.get("flagged_aps", []):
        band_label = f"{a['band']} GHz" if a.get("band") else ""
        combined_risk_items.append({
            "deviceType": "MR",
            "item": f"{a['name']} ({band_label})" if band_label else a["name"],
            "network": a.get("network", ""),
            "metric": "RF Channel Utilization",
            "value": f"{a['utilizationPercent']}%",
        })

    return {"combined_risk_items": combined_risk_items}


# ---------------------------------------------------------------------------
# 14. SWITCH UPLINK STATUS (offline / degraded speed on uplink ports)
#     NOTE: "port instability" (flapping) requires polling this same data
#     repeatedly over time and diffing snapshots - a single run can only
#     see a point-in-time status, so that specific metric is not included
#     here. Run this script on a recurring schedule if you need flapping
#     detection; the offline/degraded checks below work from a single run.
# ---------------------------------------------------------------------------
def get_switch_uplink_status(switch_data):
    offline_uplinks = []
    degraded_uplinks = []
    for switch_entry in switch_data.get("port_statuses", []):
        switch_name = switch_entry.get("name", switch_entry.get("serial"))
        for port in switch_entry.get("ports", []):
            if not port.get("isUplink"):
                continue
            status = port.get("status", "")
            speed = port.get("speed", "") or ""
            if status not in ("Connected",):
                offline_uplinks.append({
                    "switch": switch_name,
                    "serial": switch_entry.get("serial"),
                    "portId": port.get("portId"),
                    "status": status,
                })
            elif any(slow in speed for slow in NON_GIGABIT_SPEEDS):
                degraded_uplinks.append({
                    "switch": switch_name,
                    "serial": switch_entry.get("serial"),
                    "portId": port.get("portId"),
                    "speed": speed,
                })
    return {"offline_uplinks": offline_uplinks, "degraded_uplinks": degraded_uplinks}


# ---------------------------------------------------------------------------
# 15. SITES POSSIBLY NEEDING AP EXPANSION
#     Heuristic only: Meraki has no direct "needs more APs" metric. This
#     reuses the wireless connection-failure data already collected - high
#     assoc/auth/DHCP/DNS failure rates are a reasonable proxy for
#     weak-signal / overloaded-AP conditions, but this is an approximation,
#     not an official Meraki calculation.
# ---------------------------------------------------------------------------
def get_ap_expansion_candidates(wireless_data):
    return wireless_data.get("flagged_networks", [])


# ---------------------------------------------------------------------------
# 16. LICENSE STATUS
#     Org licensing can be either co-term (one shared license pool) or
#     per-device (PDL). We try the overview endpoint (works for co-term)
#     and the per-device list (works for PDL) - whichever doesn't apply
#     to this org will just 400/404 and get skipped gracefully.
# ---------------------------------------------------------------------------
def get_license_status(dashboard, org_id):
    log("Checking license status...")
    overview = call_with_retry(
        dashboard.organizations.getOrganizationLicensesOverview, org_id
    )

    per_device_licenses = call_with_retry(
        dashboard.organizations.getOrganizationLicenses, org_id, total_pages="all"
    ) or []

    unassigned_or_expired = []
    expiring_soon = []
    now = datetime.now(timezone.utc)
    for lic in per_device_licenses:
        state = lic.get("state", "")
        if state in ("unassigned", "expired"):
            unassigned_or_expired.append({
                "licenseType": lic.get("licenseType"),
                "deviceSerial": lic.get("deviceSerial"),
                "state": state,
            })
        exp = lic.get("expirationDate")
        if exp:
            exp_dt = _parse_ts(exp)
            if exp_dt and 0 <= (exp_dt - now).days <= LICENSE_EXPIRING_SOON_DAYS:
                expiring_soon.append({
                    "licenseType": lic.get("licenseType"),
                    "deviceSerial": lic.get("deviceSerial"),
                    "expirationDate": exp,
                })

    return {
        "overview": overview,
        "unassigned_or_expired": unassigned_or_expired,
        "expiring_soon": expiring_soon,
    }


# ---------------------------------------------------------------------------
# 17. SECURITY EVENTS (AMP malware blocks, IDS/IPS alerts)
# ---------------------------------------------------------------------------
def _build_attack_description(e):
    """Builds the human-readable 'what is this' detail for a security event.
    The two event shapes carry the actual description in different fields:
      - Malware/AMP events ("File Scanned" etc.): canonicalName is the real
        threat name (e.g. "PUA.Win.Dropper.Kraddare::1201").
      - IDS/IPS ("IDS Alert") events: 'message' is Snort's human-readable
        description of what the signature matched (e.g. "SERVER-OTHER Zyxel
        network device unauthenticated command injection attempt") - the
        'signature' field alone is just a cryptic GID:SID:rev number, so it's
        used as a reference tag alongside the message, never as the headline.
    """
    canonical = e.get("canonicalName")
    if canonical:
        detail = canonical
        uri = e.get("uri")
        if uri:
            detail += f" via {uri}"
        return detail

    message = e.get("message")
    if message:
        sig_ref = e.get("signature") or e.get("ruleId")
        return f"{message} (sig: {sig_ref})" if sig_ref else message

    return e.get("signature") or e.get("ruleId") or e.get("eventType") or "Unspecified"


def get_security_events(dashboard, org_id):
    log("Fetching security events (malware / IDS-IPS)...")
    events = call_with_retry(
        dashboard.appliance.getOrganizationApplianceSecurityEvents,
        org_id, total_pages="all", timespan=TIMESPAN_SECONDS
    ) or []

    malicious = [e for e in events if e.get("disposition") == "Malicious" or e.get("eventType") == "IDS Alert"]

    enriched = []
    for e in malicious[:100]:  # cap for report size
        dt = _parse_ts(e.get("ts"))
        # IDS Alert events carry a boolean 'blocked' flag instead of a
        # disposition/action string - fall back to that so the Action
        # column isn't blank for the most common event type in this list.
        if e.get("disposition") or e.get("action"):
            action_display = e.get("disposition") or e.get("action")
        elif e.get("blocked") is not None:
            action_display = "Blocked" if e.get("blocked") else "Allowed"
        else:
            action_display = ""

        enriched.append({
            "tsFormatted": format_local_str(dt) if dt else e.get("ts", ""),
            "eventType": e.get("eventType", ""),
            "srcIp": e.get("srcIp") or e.get("clientIp") or "",
            "destIp": e.get("destIp") or "",
            "destPort": e.get("destinationPort", ""),
            "protocol": e.get("protocol", ""),
            "attackDescription": _build_attack_description(e),
            "action": action_display,
            "network": e.get("network", {}).get("name", ""),
        })

    return {
        "total_events": len(events),
        "malicious_or_alert_events": enriched,
        "malicious_or_alert_count": len(malicious),
    }



# ---------------------------------------------------------------------------
# EXECUTIVE SUMMARY CALCULATION
# ---------------------------------------------------------------------------
def build_executive_summary(devices, availability_data, wan, switch, wireless, camera, sensor, firmware,
                             networks_firmware, offline_by_type, mx_utilization, device_risk, switch_uplinks,
                             ap_expansion, licensing, security):
    total_devices = len(devices)
    current = availability_data["current_snapshot"]
    offline_now = [d for d in current if d.get("status") == "offline"]
    alerting_now = [d for d in current if d.get("status") == "alerting"]

    total_downtime_seconds = sum(d["totalDowntimeSeconds"] for d in availability_data["downtime_summary"])
    possible_uptime_seconds = total_devices * TIMESPAN_SECONDS if total_devices else 1
    overall_uptime_pct = max(0, 100 - (total_downtime_seconds / possible_uptime_seconds * 100))

    return {
        "report_generated": datetime.now(timezone.utc).isoformat(),
        "timespan_days": TIMESPAN_SECONDS // 86400,
        "total_devices": total_devices,
        "devices_offline_now": len(offline_now),
        "devices_alerting_now": len(alerting_now),
        "offline_by_type": offline_by_type["offline_by_type"],
        "overall_uptime_percent": round(overall_uptime_pct, 3),
        "critical_devices_down": [d.get("name") or d.get("serial") for d in offline_now],
        "outdated_firmware_count": len(firmware["outdated_summary"]),
        "devices_with_upgrade_available": firmware["total_devices_with_upgrade_available"],
        "networks_with_pending_firmware": len(networks_firmware),
        "wan_flagged_uplinks": len(wan["flagged_uplinks"]),
        "switch_ports_flagged": len(switch["flagged_ports"]),
        "wireless_networks_flagged": len(wireless["flagged_networks"]),
        "cameras_offline_now": len(camera["offline_now"]),
        "sensors_low_battery": len(sensor["low_battery"]),
        "mx_near_capacity": len(mx_utilization["flagged_near_capacity"]),
        "devices_near_capacity_total": len(device_risk["combined_risk_items"]),
        "switch_uplinks_offline": len(switch_uplinks["offline_uplinks"]),
        "switch_uplinks_degraded": len(switch_uplinks["degraded_uplinks"]),
        "sites_possible_ap_expansion": len(ap_expansion),
        "licenses_unassigned_or_expired": len(licensing["unassigned_or_expired"]),
        "licenses_expiring_soon": len(licensing["expiring_soon"]),
        "security_malicious_or_alert_count": security["malicious_or_alert_count"],
    }


# ---------------------------------------------------------------------------
# OUTPUT: JSON
# ---------------------------------------------------------------------------
def write_json(report):
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, default=str)
    log(f"Raw telemetry JSON saved: {JSON_OUT}")


# ---------------------------------------------------------------------------
# OUTPUT: CSV SUMMARY (flat, one row per finding)
# ---------------------------------------------------------------------------
def write_csv(report):
    rows = []
    exec_summary = report["executive_summary"]
    rows.append(["Category", "Item", "Detail", "Value"])
    rows.append(["Executive", "Overall Uptime %", "", exec_summary["overall_uptime_percent"]])
    rows.append(["Executive", "Devices Offline Now", "", exec_summary["devices_offline_now"]])
    rows.append(["Executive", "Devices Alerting Now", "", exec_summary["devices_alerting_now"]])
    rows.append(["Executive", "Outdated Firmware Count", "", exec_summary["outdated_firmware_count"]])
    rows.append(["Executive", "Devices w/ Upgrade Available", "", exec_summary["devices_with_upgrade_available"]])

    for k, v in exec_summary["offline_by_type"].items():
        rows.append(["Offline by Type (Count)", k, "", v])

    for d in report["offline_by_type"]["offline_devices_detail"]:
        rows.append(["Offline Device Detail", f"{d['deviceTypeLabel']} - {d['name']}", d["network"],
                      f"Serial: {d['serial']} | {d['lastSeen']} | Downtime: {d['downtime']}"])

    for d in report["availability"]["downtime_summary"][:50]:
        rows.append(["Outages", d["name"], d["network"], f"{d['outageCount']} outages / {d['totalDowntimeHours']}h down"])

    for u in report["wan"]["flagged_uplinks"]:
        rows.append(["WAN", u.get("network"), u.get("uplink"), f"loss={u.get('lossPercent')}% latency={u.get('latencyMs')}ms"])

    for s in report["sensor"]["low_battery"]:
        rows.append(["Sensor", s["serial"], s["network"], f"battery {s['batteryPercent']}%"])

    for f_ in report["firmware"]["outdated_summary"]:
        rows.append(["Firmware", f_["network"], f_["productType"],
                      f"{f_['currentVersion']} -> {f_['availableVersion']} ({f_['affectedDeviceCount']} device(s))"])

    for n in report["networks_with_firmware_upgrades"]:
        rows.append(["Firmware Upgrades Available (by Network)", n["network"],
                      f"{n['totalDevicesAffected']} device(s) affected",
                      f"{n['productsSummary']} | Scheduled: {n['scheduledFor']}"])

    for a in report["mx_utilization"]["all_appliances"]:
        pct_display = f"{a['utilizationPercent']}%" if a["utilizationPercent"] is not None else "N/A"
        rows.append(["MX Utilization", f"{a['name']} ({a['model']})", a["network"],
                      f"{pct_display}{' - NEAR CAPACITY' if a['nearCapacity'] else ''}"])

    for item in report["device_health_risk_signals"]["combined_risk_items"]:
        rows.append([f"Utilization Risk ({item['deviceType']})", item["item"], item["network"],
                      f"{item['metric']}: {item['value']}"])

    for u in report["switch_uplink_status"]["offline_uplinks"]:
        rows.append(["Switch Uplink", u["switch"], f"port {u['portId']}", f"OFFLINE ({u['status']})"])
    for u in report["switch_uplink_status"]["degraded_uplinks"]:
        rows.append(["Switch Uplink", u["switch"], f"port {u['portId']}", f"Degraded speed: {u['speed']}"])

    for w in report["ap_expansion_candidates"]:
        rows.append(["AP Expansion Candidate", w["network"], "", f"{w['failurePercent']}% connection failures"])

    for l in report["licensing"]["unassigned_or_expired"]:
        rows.append(["License", l["licenseType"], l["deviceSerial"], l["state"]])
    for l in report["licensing"]["expiring_soon"]:
        rows.append(["License Expiring Soon", l["licenseType"], l["deviceSerial"], l["expirationDate"]])

    for e in report["security_events"]["malicious_or_alert_events"]:
        rows.append(["Security Event", e["tsFormatted"], f"{e['srcIp']} -> {e['destIp']}",
                      f"{e['eventType']}: {e['attackDescription']} ({e['action']})"])

    with open(CSV_OUT, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    log(f"CSV summary saved: {CSV_OUT}")


# ---------------------------------------------------------------------------
# OUTPUT: HTML EXECUTIVE REPORT
# ---------------------------------------------------------------------------
def write_html(report, org_name):
    exec_ = report["executive_summary"]

    def status_color(pct):
        if pct >= 99.5:
            return "#1a7f37"
        if pct >= 98:
            return "#9a6700"
        return "#cf222e"

    uptime_color = status_color(exec_["overall_uptime_percent"])

    outage_rows = ""
    for d in report["availability"]["downtime_summary"][:30]:
        outage_rows += (f"<tr><td>{d['name']}</td><td>{d['network']}</td>"
                         f"<td>{d['productType']}</td><td>{d['outageCount']}</td>"
                         f"<td>{d['totalDowntimeHours']}</td></tr>")
    if not outage_rows:
        outage_rows = "<tr><td colspan='5' class='empty'>No outages recorded in this window.</td></tr>"

    wan_rows = ""
    for u in report["wan"]["flagged_uplinks"]:
        wan_rows += (f"<tr><td>{u.get('network')}</td><td>{u.get('uplink')}</td>"
                      f"<td>{u.get('lossPercent')}</td><td>{u.get('latencyMs')}</td></tr>")
    if not wan_rows:
        wan_rows = "<tr><td colspan='4' class='empty'>No WAN uplinks exceeded loss/latency thresholds.</td></tr>"

    sensor_rows = ""
    for s in report["sensor"]["low_battery"]:
        sensor_rows += f"<tr><td>{s['serial']}</td><td>{s['network']}</td><td>{s['batteryPercent']}%</td></tr>"
    if not sensor_rows:
        sensor_rows = "<tr><td colspan='3' class='empty'>No sensors below the low-battery threshold.</td></tr>"

    firmware_rows = ""
    for f_ in report["firmware"]["outdated_summary"]:
        firmware_rows += (f"<tr><td>{f_['network']}</td><td>{f_['productType']}</td>"
                           f"<td>{f_['currentVersion']}</td><td>{f_['availableVersion']}</td>"
                           f"<td>{f_['affectedDeviceCount']}</td></tr>")
    if not firmware_rows:
        firmware_rows = "<tr><td colspan='5' class='empty'>All devices on latest recommended firmware.</td></tr>"

    network_firmware_rows = ""
    for n in report["networks_with_firmware_upgrades"]:
        network_firmware_rows += (f"<tr><td>{n['network']}</td><td>{n['productsSummary']}</td>"
                                   f"<td>{n['totalDevicesAffected']}</td><td>{n['scheduledFor']}</td></tr>")
    if not network_firmware_rows:
        network_firmware_rows = "<tr><td colspan='4' class='empty'>No networks have a firmware upgrade available.</td></tr>"


    offline_detail_rows = ""
    for d in report["offline_by_type"]["offline_devices_detail"]:
        offline_detail_rows += (f"<tr><td>{d['deviceTypeLabel']}</td><td>{d['name']}</td>"
                                 f"<td>{d['network']}</td><td>{d['serial']}</td>"
                                 f"<td>{d['lastSeen']}</td><td>{d['downtime']}</td></tr>")
    if not offline_detail_rows:
        offline_detail_rows = "<tr><td colspan='6' class='empty'>No devices currently offline.</td></tr>"

    mx_util_rows = ""
    for a in report["mx_utilization"]["all_appliances"]:
        pct_display = f"{a['utilizationPercent']}%" if a["utilizationPercent"] is not None else "N/A"
        row_color = "#cf222e" if a["nearCapacity"] else "#1f2328"
        flag_display = "Near Capacity" if a["nearCapacity"] else ""
        mx_util_rows += (f"<tr><td>{a['network']}</td><td>{a['name']}</td><td>{a['model']}</td>"
                          f"<td style='color:{row_color};font-weight:{'600' if a['nearCapacity'] else '400'}'>{pct_display}</td>"
                          f"<td>{flag_display}</td></tr>")
    if not mx_util_rows:
        mx_util_rows = "<tr><td colspan='5' class='empty'>No MX utilization data returned for this org.</td></tr>"

    device_util_rows = ""
    for item in report["device_health_risk_signals"]["combined_risk_items"]:
        device_util_rows += (f"<tr><td>{item['deviceType']}</td><td>{item['item']}</td>"
                              f"<td>{item['network']}</td><td>{item['metric']}</td>"
                              f"<td>{item['value']}</td></tr>")
    if not device_util_rows:
        device_util_rows = "<tr><td colspan='5' class='empty'>No devices flagged near capacity or at risk.</td></tr>"

    switch_uplink_rows = ""
    for u in report["switch_uplink_status"]["offline_uplinks"]:
        switch_uplink_rows += f"<tr><td>{u['switch']}</td><td>{u['portId']}</td><td>Offline</td><td>{u['status']}</td></tr>"
    for u in report["switch_uplink_status"]["degraded_uplinks"]:
        switch_uplink_rows += f"<tr><td>{u['switch']}</td><td>{u['portId']}</td><td>Degraded Speed</td><td>{u['speed']}</td></tr>"
    if not switch_uplink_rows:
        switch_uplink_rows = "<tr><td colspan='4' class='empty'>All switch uplink ports online and at full speed.</td></tr>"

    ap_expansion_rows = ""
    for w in report["ap_expansion_candidates"]:
        ap_expansion_rows += f"<tr><td>{w['network']}</td><td>{w['failurePercent']}%</td></tr>"
    if not ap_expansion_rows:
        ap_expansion_rows = "<tr><td colspan='2' class='empty'>No sites flagged for possible AP expansion.</td></tr>"

    license_rows = ""
    for l in report["licensing"]["unassigned_or_expired"]:
        license_rows += f"<tr><td>Unassigned/Expired</td><td>{l['licenseType']} - serial {l['deviceSerial']} ({l['state']})</td></tr>"
    for l in report["licensing"]["expiring_soon"]:
        license_rows += f"<tr><td>Expiring Soon</td><td>{l['licenseType']} - serial {l['deviceSerial']} (expires {l['expirationDate']})</td></tr>"
    if not license_rows:
        license_rows = "<tr><td colspan='2' class='empty'>No unassigned, expired, or soon-to-expire licenses found.</td></tr>"

    security_rows = ""
    for e in report["security_events"]["malicious_or_alert_events"][:50]:
        security_rows += (f"<tr><td>{e['tsFormatted']}</td><td>{e['eventType']}</td>"
                           f"<td>{e['srcIp']}</td><td>{e['destIp']}</td><td>{e['destPort']}</td>"
                           f"<td>{e['attackDescription']}</td><td>{e['action']}</td></tr>")
    if not security_rows:
        security_rows = "<tr><td colspan='7' class='empty'>No malware or IDS/IPS alerts detected in this window.</td></tr>"

    # Unique id for THIS run, embedded in the page and in latest.json.
    # The in-page poller compares these to detect a newer run.
    generated_iso = datetime.now(timezone.utc).isoformat()

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Network Health Report - {org_name}</title>
<meta name="report-generated" content="{generated_iso}">
<style>
  html {{ scroll-behavior: smooth; }}
  body {{ font-family: -apple-system, Segoe UI, Helvetica, Arial, sans-serif; background:#f6f8fa; color:#1f2328; margin:0; padding:40px; }}
  .container {{ max-width: 1100px; margin: 0 auto; }}
  h1 {{ font-size: 26px; margin-bottom:4px; }}
  .subtitle {{ color:#57606a; margin-bottom:28px; }}
  .cards {{ display:flex; gap:16px; flex-wrap:wrap; margin-bottom:36px; }}
  a.card-link {{ text-decoration:none; color:inherit; flex:1; min-width:180px; }}
  .card {{ background:white; border:1px solid #d0d7de; border-radius:8px; padding:18px 22px; height:100%; box-sizing:border-box; transition:box-shadow .15s ease, transform .15s ease; }}
  a.card-link:hover .card {{ box-shadow:0 2px 8px rgba(0,0,0,0.08); transform:translateY(-1px); cursor:pointer; }}
  .card.static {{ cursor:default; }}
  .card .label {{ font-size:12px; text-transform:uppercase; color:#57606a; letter-spacing:.04em; }}
  .card .value {{ font-size:28px; font-weight:600; margin-top:4px; }}
  section {{ background:white; border:1px solid #d0d7de; border-radius:8px; padding:22px 24px; margin-bottom:24px; scroll-margin-top:20px; }}
  section h2 {{ font-size:17px; margin-top:0; border-bottom:1px solid #d0d7de; padding-bottom:10px; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ text-align:left; color:#57606a; font-weight:600; padding:8px 10px; border-bottom:1px solid #d0d7de; }}
  td {{ padding:8px 10px; border-bottom:1px solid #eaeef2; }}
  .empty {{ color:#57606a; font-style:italic; text-align:center; padding:16px; }}
  footer {{ color:#8b949e; font-size:12px; text-align:center; margin-top:30px; }}
</style>
</head>
<body>
<div class="container">
  <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;">
    <div>
      <h1>Network Health Report</h1>
      <div class="subtitle">{org_name} &middot; Last {exec_['timespan_days']} days &middot; Generated {exec_['report_generated']}</div>
    </div>
    <div style="text-align:right;">
      <button id="refresh-btn" onclick="checkForNewerReport(true)"
        style="background:#1f6feb;color:white;border:none;border-radius:6px;padding:9px 16px;font-size:13px;font-weight:600;cursor:pointer;">
        &#8635; Refresh
      </button>
      <div id="refresh-status" style="font-size:11px;color:#57606a;margin-top:6px;"></div>
    </div>
  </div>
  <div id="update-banner" style="display:none;background:#ddf4ff;border:1px solid #54aeff;border-radius:6px;padding:10px 16px;margin-bottom:20px;font-size:13px;">
    A newer report is available. <a href="#" onclick="location.reload();return false;" style="font-weight:600;">Reload to view it</a>.
  </div>

  <div class="cards">
    <a class="card-link" href="#offline-devices"><div class="card"><div class="label">Overall Uptime</div><div class="value" style="color:{uptime_color}">{exec_['overall_uptime_percent']}%</div></div></a>
    <div class="card-link"><div class="card static"><div class="label">Total Devices</div><div class="value">{exec_['total_devices']}</div></div></div>
    <a class="card-link" href="#offline-devices"><div class="card"><div class="label">Offline Now</div><div class="value" style="color:#cf222e">{exec_['devices_offline_now']}</div></div></a>
    <div class="card-link"><div class="card static"><div class="label">Alerting Now</div><div class="value" style="color:#9a6700">{exec_['devices_alerting_now']}</div></div></div>
    <a class="card-link" href="#firmware-compliance"><div class="card"><div class="label">Outdated Firmware</div><div class="value">{exec_['outdated_firmware_count']}</div></div></a>
    <a class="card-link" href="#firmware-compliance"><div class="card"><div class="label">Devices w/ Upgrade Available</div><div class="value">{exec_['devices_with_upgrade_available']}</div></div></a>
    <a class="card-link" href="#networks-firmware-upgrades"><div class="card"><div class="label">Networks w/ Firmware Pending</div><div class="value" style="color:{'#9a6700' if exec_['networks_with_pending_firmware'] else '#1a7f37'}">{exec_['networks_with_pending_firmware']}</div></div></a>
    <a class="card-link" href="#wan-performance"><div class="card"><div class="label">WAN Uplinks Flagged</div><div class="value" style="color:{'#cf222e' if exec_['wan_flagged_uplinks'] else '#1a7f37'}">{exec_['wan_flagged_uplinks']}</div></div></a>
    <a class="card-link" href="#mx-utilization"><div class="card"><div class="label">MX Near Capacity</div><div class="value" style="color:{'#cf222e' if exec_['mx_near_capacity'] else '#1a7f37'}">{exec_['mx_near_capacity']}</div></div></a>
    <a class="card-link" href="#device-utilization"><div class="card"><div class="label">MS/MR Risk Signals</div><div class="value" style="color:{'#9a6700' if exec_['devices_near_capacity_total'] else '#1a7f37'}">{exec_['devices_near_capacity_total']}</div></div></a>
    <a class="card-link" href="#ap-expansion"><div class="card"><div class="label">Sites - Possible AP Expansion</div><div class="value" style="color:{'#9a6700' if exec_['sites_possible_ap_expansion'] else '#1a7f37'}">{exec_['sites_possible_ap_expansion']}</div></div></a>
    <a class="card-link" href="#switch-uplink-status"><div class="card"><div class="label">Switch Uplinks Offline</div><div class="value" style="color:{'#cf222e' if exec_['switch_uplinks_offline'] else '#1a7f37'}">{exec_['switch_uplinks_offline']}</div></div></a>
    <a class="card-link" href="#switch-uplink-status"><div class="card"><div class="label">Switch Uplinks Degraded</div><div class="value" style="color:{'#9a6700' if exec_['switch_uplinks_degraded'] else '#1a7f37'}">{exec_['switch_uplinks_degraded']}</div></div></a>
    <a class="card-link" href="#license-status"><div class="card"><div class="label">Licenses Unassigned/Expired</div><div class="value" style="color:{'#cf222e' if exec_['licenses_unassigned_or_expired'] else '#1a7f37'}">{exec_['licenses_unassigned_or_expired']}</div></div></a>
    <a class="card-link" href="#license-status"><div class="card"><div class="label">Licenses Expiring Soon</div><div class="value" style="color:{'#9a6700' if exec_['licenses_expiring_soon'] else '#1a7f37'}">{exec_['licenses_expiring_soon']}</div></div></a>
    <a class="card-link" href="#security-events"><div class="card"><div class="label">Security Events (Malware/IDS)</div><div class="value" style="color:{'#cf222e' if exec_['security_malicious_or_alert_count'] else '#1a7f37'}">{exec_['security_malicious_or_alert_count']}</div></div></a>
  </div>

  <section id="offline-devices">
    <h2>Offline Devices by Type</h2>
    <table><tr><th>Device Type</th><th>Device Name</th><th>Network</th><th>Serial #</th><th>Last Seen</th><th>Downtime</th></tr>
    {offline_detail_rows}
    </table>
  </section>

  <section id="device-outages">
    <h2>Device Outages (Top 30 by downtime, last {exec_['timespan_days']} days)</h2>
    <table><tr><th>Device</th><th>Network</th><th>Type</th><th>Outage Count</th><th>Total Downtime (hrs)</th></tr>
    {outage_rows}
    </table>
  </section>

  <section id="wan-performance">
    <h2>WAN &amp; ISP Performance</h2>
    <table><tr><th>Network</th><th>Uplink</th><th>Loss %</th><th>Latency (ms)</th></tr>
    {wan_rows}
    </table>
  </section>

  <section id="mx-utilization">
    <h2>MX (Security Appliance) Utilization</h2>
    <p style="color:#57606a;font-size:13px;margin-top:-6px;">Meraki-reported utilization: current throughput vs. each model's rated maximum throughput. Rows in red are at or above the {APPLIANCE_UTILIZATION_THRESHOLD_PCT}% threshold - a useful signal for whether the site's MX model needs to be upgraded. Note: Meraki's API only returns the top 10 appliances by utilization org-wide, so appliances outside the top 10 (i.e. the least-utilized ones) won't appear here.</p>
    <table><tr><th>Network</th><th>Appliance</th><th>Model</th><th>Utilization %</th><th>Flag</th></tr>
    {mx_util_rows}
    </table>
  </section>

  <section id="device-utilization">
    <h2>Device Health &amp; Risk Signals - MS &amp; MR</h2>
    <p style="color:#57606a;font-size:13px;margin-top:-6px;">Real capacity metrics, not proxies: MS rows are measured port traffic against that port's rated link speed; MR rows are RF channel utilization (spectrum saturation) at that AP/band. Camera and sensor issues aren't capacity signals, so they're covered under Offline Devices and Environmental Sensors instead.</p>
    <table><tr><th>Device Type</th><th>Item</th><th>Network</th><th>Metric</th><th>Value</th></tr>
    {device_util_rows}
    </table>
  </section>

  <section id="ap-expansion">
    <h2>Sites - Possible AP Expansion Candidates</h2>
    <p style="color:#57606a;font-size:13px;margin-top:-6px;">Heuristic based on wireless connection-failure rate (association/auth/DHCP/DNS failures) - not an official Meraki metric, but a reasonable proxy for weak-signal or overloaded-AP conditions.</p>
    <table><tr><th>Network</th><th>Connection Failure %</th></tr>
    {ap_expansion_rows}
    </table>
  </section>

  <section id="switch-uplink-status">
    <h2>Switch Uplink Status</h2>
    <p style="color:#57606a;font-size:13px;margin-top:-6px;">Offline and speed-degraded uplinks from this run's snapshot. Port "instability" (flapping) requires recurring runs of this script to detect and is not shown here.</p>
    <table><tr><th>Switch</th><th>Port</th><th>Issue</th><th>Detail</th></tr>
    {switch_uplink_rows}
    </table>
  </section>

  <section>
    <h2>Environmental Sensors - Low Battery</h2>
    <table><tr><th>Serial</th><th>Network</th><th>Battery</th></tr>
    {sensor_rows}
    </table>
  </section>

  <section id="networks-firmware-upgrades">
    <h2>Networks with Firmware Upgrades Available</h2>
    <p style="color:#57606a;font-size:13px;margin-top:-6px;">One row per network, rolled up from the per-product table below - a quick view of which sites have anything pending, with the earliest scheduled upgrade time if Meraki has one set.</p>
    <table><tr><th>Network</th><th>Products w/ Upgrade Available</th><th>Devices Affected</th><th>Earliest Scheduled</th></tr>
    {network_firmware_rows}
    </table>
  </section>

  <section id="firmware-compliance">
    <h2>Firmware Compliance - Devices Pending Upgrade</h2>
    <table><tr><th>Network</th><th>Product</th><th>Current</th><th>Available</th><th>Devices Affected</th></tr>
    {firmware_rows}
    </table>
  </section>

  <section id="license-status">
    <h2>License Status</h2>
    <table><tr><th>Category</th><th>Detail</th></tr>
    {license_rows}
    </table>
  </section>

  <section id="security-events">
    <h2>Security Events - Malware / IDS-IPS Alerts (last {exec_['timespan_days']} days)</h2>
    <table><tr><th>Timestamp</th><th>Event Type</th><th>Source IP</th><th>Destination IP</th><th>Dest Port</th><th>Attack / Detail</th><th>Action</th></tr>
    {security_rows}
    </table>
  </section>

  <footer>Generated automatically via the Meraki Dashboard API. Data reflects a {exec_['timespan_days']}-day lookback window.</footer>
</div>
<script>
  // This page's own generation time, embedded server-side by write_html().
  const THIS_REPORT_GENERATED = document.querySelector('meta[name="report-generated"]').content;

  // "Refresh" does NOT call the Meraki API from the browser - that would mean
  // shipping your API key to every teammate's browser, which is unsafe and
  // won't work anyway (Meraki's API blocks direct browser/CORS requests).
  // Instead it checks latest.json (written by the same script run that
  // produced this page) to see whether a newer run has been generated and
  // published (e.g. by a scheduled job - see the README note in the script).
  async function checkForNewerReport(userInitiated) {{
    const statusEl = document.getElementById('refresh-status');
    const banner = document.getElementById('update-banner');
    try {{
      if (userInitiated) statusEl.textContent = 'Checking...';
      const res = await fetch('latest.json', {{ cache: 'no-store' }});
      if (!res.ok) throw new Error('no manifest');
      const data = await res.json();
      if (data.generated && data.generated !== THIS_REPORT_GENERATED) {{
        banner.style.display = 'block';
        statusEl.textContent = 'Newer report found.';
      }} else {{
        statusEl.textContent = userInitiated ? 'You have the latest report.' : '';
      }}
    }} catch (e) {{
      // latest.json isn't reachable (e.g. opened as a local double-clicked
      // file, or the manifest hasn't been published to this host) - fail
      // quietly rather than alarming the user.
      if (userInitiated) statusEl.textContent = 'Could not check for updates.';
    }}
  }}

  // Passive check on load, then every 5 minutes, so the banner can appear
  // on its own for anyone who leaves the tab open.
  checkForNewerReport(false);
  setInterval(() => checkForNewerReport(false), 5 * 60 * 1000);
</script>
</body>
</html>"""

    with open(HTML_OUT, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"Executive HTML report saved: {HTML_OUT}")

    # Stable copy teammates can bookmark - a fixed URL/filename that always
    # holds the most recent run, instead of a new timestamped filename every time.
    with open(HTML_LATEST, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"Stable 'latest' HTML copy saved: {HTML_LATEST}")

    # Tiny manifest the page's JS polls to detect a newer run.
    with open(MANIFEST_OUT, "w", encoding="utf-8") as f:
        json.dump({"generated": generated_iso, "organization": org_name}, f)
    log(f"Manifest saved: {MANIFEST_OUT}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    if not API_KEY or API_KEY == "PASTE_YOUR_MERAKI_API_KEY_HERE":
        log("ERROR: Edit the API_KEY variable near the top of this script and paste in your real Meraki API key.")
        sys.exit(1)

    dashboard = meraki.DashboardAPI(
        api_key=API_KEY,
        suppress_logging=True,
        wait_on_rate_limit=True,   # SDK-level 429 handling
        maximum_retries=5,
        nginx_429_retry_wait_time=5,
    )

    org_id_to_use = ORG_ID if ORG_ID and ORG_ID != "PASTE_YOUR_ORG_ID_HERE" else None
    org = choose_organization(dashboard, org_id_to_use)
    org_id = org["id"]
    org_name = org["name"]
    log(f"Generating report for: {org_name} (ID: {org_id})")

    devices, devices_by_type = get_all_devices(dashboard, org_id)

    networks = call_with_retry(
        dashboard.organizations.getOrganizationNetworks, org_id, total_pages="all"
    ) or []
    network_id_to_name = {n["id"]: n.get("name", n["id"]) for n in networks}

    availability_data = get_availability_data(dashboard, org_id)
    wan_data = get_wan_health(dashboard, org_id)
    switch_data = get_switch_health(dashboard, org_id, devices_by_type.get("switch", []))
    wireless_data = get_wireless_health(dashboard, org_id, networks)
    camera_data = get_camera_health(availability_data)
    sensor_data = get_sensor_health(dashboard, org_id, devices_by_type.get("sensor", []))
    firmware_data = get_firmware_compliance(dashboard, org_id, devices, networks)
    networks_firmware_data = get_networks_with_firmware_upgrades(firmware_data)

    offline_by_type_data = get_offline_by_type(availability_data, network_id_to_name)
    ap_channel_util_data = get_ap_channel_utilization(
        dashboard, org_id, devices_by_type.get("wireless", []), network_id_to_name
    )
    mx_utilization_data = get_mx_utilization_overview(dashboard, org_id)
    device_risk_data = get_device_health_risk_signals(switch_data, ap_channel_util_data)
    switch_uplink_data = get_switch_uplink_status(switch_data)
    ap_expansion_data = get_ap_expansion_candidates(wireless_data)
    license_data = get_license_status(dashboard, org_id)
    security_data = get_security_events(dashboard, org_id)

    executive_summary = build_executive_summary(
        devices, availability_data, wan_data, switch_data,
        wireless_data, camera_data, sensor_data, firmware_data, networks_firmware_data,
        offline_by_type_data, mx_utilization_data, device_risk_data, switch_uplink_data,
        ap_expansion_data, license_data, security_data
    )

    report = {
        "organization": {"id": org_id, "name": org_name},
        "executive_summary": executive_summary,
        "devices_by_type": {k: len(v) for k, v in devices_by_type.items()},
        "availability": availability_data,
        "wan": wan_data,
        "switch": switch_data,
        "wireless": wireless_data,
        "camera": camera_data,
        "sensor": sensor_data,
        "firmware": firmware_data,
        "networks_with_firmware_upgrades": networks_firmware_data,
        "offline_by_type": offline_by_type_data,
        "mx_utilization": mx_utilization_data,
        "ap_channel_utilization": ap_channel_util_data,
        "device_health_risk_signals": device_risk_data,
        "switch_uplink_status": switch_uplink_data,
        "ap_expansion_candidates": ap_expansion_data,
        "licensing": license_data,
        "security_events": security_data,
    }

    write_json(report)
    write_csv(report)
    write_html(report, org_name)

    print("\n" + "=" * 60)
    print(f"  Overall Uptime:        {executive_summary['overall_uptime_percent']}%")
    print(f"  Devices Offline Now:   {executive_summary['devices_offline_now']}")
    print(f"  Devices Alerting Now:  {executive_summary['devices_alerting_now']}")
    print(f"  Outdated Firmware:     {executive_summary['outdated_firmware_count']}")
    print("=" * 60)
    print(f"\nReports saved:\n  {JSON_OUT}\n  {CSV_OUT}\n  {HTML_OUT}\n")


if __name__ == "__main__":
    main()
