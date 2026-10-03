"""
cw_exporter.py - CloudWatch Logs Fetcher & Excel Report Generator.
Uses pure Python standard libraries (zero external dependencies).
Interacts with AWS CLI to retrieve CloudWatch log groups, streams, and events,
then delegates to xlsx_builder to create a beautifully styled Excel workbook.
"""

import sys
import os
import re
import json
import subprocess
import argparse
import datetime
import time

# Ensure lib directory is on sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from xlsx_builder import (
    XlsxWorkbook,
    STYLE_TITLE,
    STYLE_META_LABEL,
    STYLE_META_VALUE,
    STYLE_HEADER_CENTER,
    STYLE_HEADER_LEFT,
    STYLE_STREAM_BANNER,
    STYLE_DATA_CENTER,
    STYLE_DATA_CENTER_ZEBRA,
    STYLE_DATA_LEFT,
    STYLE_DATA_LEFT_ZEBRA,
    STYLE_DATA_MSG,
    STYLE_DATA_MSG_ZEBRA,
    STYLE_HIGHLIGHT_CENTER,
    STYLE_HIGHLIGHT_LEFT,
    STYLE_HIGHLIGHT_MSG,
    STYLE_SUMMARY_HEADER_CENTER,
    STYLE_SUMMARY_HEADER_LEFT,
    STYLE_SECTION_HEADER
)

# JST timezone helper (UTC+9)
JST = datetime.timezone(datetime.timedelta(hours=9), 'JST')

def parse_time_to_epoch_ms(time_str, default_now=False):
    """
    Parses a variety of time formats to UNIX epoch milliseconds:
    - Relative: '1h', '30m', '24h', '7d'
    - Absolute: 'YYYY-MM-DD HH:MM:SS', 'YYYY-MM-DDTHH:MM:SS', 'YYYY-MM-DD'
    - Epoch: '1728000000000' (ms) or '1728000000' (sec)
    """
    if not time_str:
        if default_now:
            now_dt = datetime.datetime.now(JST)
            return int(now_dt.timestamp() * 1000)
        return None

    time_str = time_str.strip()

    # Numeric epoch
    if time_str.isdigit():
        val = int(time_str)
        if val < 10000000000:  # Seconds
            return val * 1000
        return val

    # Relative time: 10m, 1h, 24h, 7d
    m_rel = re.match(r'^(\d+)\s*([mhdMHD])$', time_str)
    if m_rel:
        num = int(m_rel.group(1))
        unit = m_rel.group(2).lower()
        now_dt = datetime.datetime.now(JST)
        if unit == 'm':
            delta = datetime.timedelta(minutes=num)
        elif unit == 'h':
            delta = datetime.timedelta(hours=num)
        elif unit == 'd':
            delta = datetime.timedelta(days=num)
        else:
            delta = datetime.timedelta(hours=1)
        target_dt = now_dt - delta
        return int(target_dt.timestamp() * 1000)

    # Absolute formats
    clean_str = time_str.replace('/', '-')
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%dT%H:%M:%S',
        '%Y-%m-%d %H:%M:%S%z',
        '%Y-%m-%dT%H:%M:%S%z',
        '%Y-%m-%d'
    ]

    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(clean_str, fmt)
            if dt.tzinfo is None:
                # Assume JST by default for local time
                dt = dt.replace(tzinfo=JST)
            return int(dt.timestamp() * 1000)
        except ValueError:
            pass

    raise ValueError(f"Unsupported time format: '{time_str}'. Use 'YYYY-MM-DD HH:MM:SS' or '1h', '24h'.")

def epoch_ms_to_jst_str(ms):
    """Converts epoch milliseconds to JST formatted string: YYYY-MM-DD HH:MM:SS.fff"""
    if not ms:
        return ""
    dt = datetime.datetime.fromtimestamp(ms / 1000.0, tz=JST)
    return dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

def run_aws_command(cmd_args):
    """Runs an AWS CLI command with JSON output, returning parsed JSON."""
    try:
        proc = subprocess.run(
            cmd_args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True
        )
        return json.loads(proc.stdout)
    except subprocess.CalledProcessError as e:
        err_msg = e.stderr.strip() if e.stderr else str(e)
        raise RuntimeError(f"AWS CLI command failed: {' '.join(cmd_args)}\nError: {err_msg}")
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Failed to parse AWS CLI JSON output: {e}")

def get_log_groups(prefix=None, specific_groups=None, region=None, profile=None):
    """Retrieves all CloudWatch log group names matching filters."""
    if specific_groups:
        return [g.strip() for g in specific_groups if g.strip()]

    base_cmd = ["aws", "logs", "describe-log-groups", "--output", "json"]
    if region:
        base_cmd.extend(["--region", region])
    if profile:
        base_cmd.extend(["--profile", profile])
    if prefix:
        base_cmd.extend(["--log-group-name-prefix", prefix])

    log_groups = []
    next_token = None

    while True:
        cmd = list(base_cmd)
        if next_token:
            cmd.extend(["--next-token", next_token])

        data = run_aws_command(cmd)
        for g in data.get("logGroups", []):
            name = g.get("logGroupName")
            if name:
                log_groups.append(name)

        next_token = data.get("nextToken")
        if not next_token:
            break

    return log_groups

def get_log_events_for_group(log_group, start_ms, end_ms, region=None, profile=None):
    """
    Fetches all log events for a given log group within the time window.
    Returns events grouped by logStreamName.
    """
    base_cmd = [
        "aws", "logs", "filter-log-events",
        "--log-group-name", log_group,
        "--output", "json"
    ]
    if start_ms is not None:
        base_cmd.extend(["--start-time", str(start_ms)])
    if end_ms is not None:
        base_cmd.extend(["--end-time", str(end_ms)])
    if region:
        base_cmd.extend(["--region", region])
    if profile:
        base_cmd.extend(["--profile", profile])

    # stream_name -> list of event dicts
    stream_events = {}
    next_token = None

    while True:
        cmd = list(base_cmd)
        if next_token:
            cmd.extend(["--next-token", next_token])

        data = run_aws_command(cmd)
        events = data.get("events", [])
        for ev in events:
            stream_name = ev.get("logStreamName", "unknown-stream")
            if stream_name not in stream_events:
                stream_events[stream_name] = []
            stream_events[stream_name].append({
                "timestamp": ev.get("timestamp", 0),
                "eventId": ev.get("eventId", ""),
                "message": ev.get("message", "")
            })

        next_token = data.get("nextToken")
        if not next_token:
            break

    # Sort each stream's events by timestamp
    for s_name in stream_events:
        stream_events[s_name].sort(key=lambda x: x["timestamp"])

    return stream_events

def build_excel_report(output_file, report_meta, groups_data, highlight_words):
    """
    Builds the complete Excel workbook using xlsx_builder.
    - Summary sheet with overview statistics and table of log groups.
    - A dedicated sheet per log group with stream blocks and highlighted keywords.
    """
    wb = XlsxWorkbook()

    # Precompile highlight regex (case-insensitive)
    hl_patterns = []
    if highlight_words:
        for w in highlight_words:
            w_clean = w.strip()
            if w_clean:
                hl_patterns.append(re.compile(re.escape(w_clean), re.IGNORECASE))

    def has_highlight(msg):
        if not msg or not hl_patterns:
            return False
        return any(p.search(msg) for p in hl_patterns)

    # Pre-create worksheets for each log group to know their final sanitized sheet names
    group_worksheets = {}
    for g in groups_data:
        ws = wb.add_worksheet(g["name"], freeze_row=4)
        g["sheet_name"] = ws.name
        group_worksheets[g["name"]] = ws

    # -------------------------------------------------------------
    # 1. Summary Sheet
    # -------------------------------------------------------------
    ws_sum = wb.add_worksheet("Summary", freeze_row=8)
    # Move Summary to the first position in sheets list so it opens first
    wb.sheets.insert(0, wb.sheets.pop())
    ws_sum.sheet_id = 1
    # Re-assign sheet_ids sequentially
    for idx, s in enumerate(wb.sheets, 1):
        s.sheet_id = idx

    ws_sum.set_row_height(1, 28)
    ws_sum.set_cell(1, 1, "CloudWatch Logs Export Report", STYLE_TITLE)
    ws_sum.merge_cells("A1:G1")

    # Metadata block
    ws_sum.set_row_height(2, 20)
    ws_sum.set_cell(2, 1, "Report Date (JST):", STYLE_META_LABEL)
    ws_sum.set_cell(2, 2, report_meta.get("generated_at", ""), STYLE_META_VALUE)
    ws_sum.set_cell(2, 3, "Total Log Groups:", STYLE_META_LABEL)
    ws_sum.set_cell(2, 4, len(groups_data), STYLE_META_VALUE)

    ws_sum.set_row_height(3, 20)
    ws_sum.set_cell(3, 1, "Time Range (JST):", STYLE_META_LABEL)
    time_range_str = f"{report_meta.get('start_time_str', 'N/A')}  ~  {report_meta.get('end_time_str', 'N/A')}"
    ws_sum.set_cell(3, 2, time_range_str, STYLE_META_VALUE)
    ws_sum.set_cell(3, 3, "Total Events:", STYLE_META_LABEL)
    total_events = sum(g["total_events"] for g in groups_data)
    ws_sum.set_cell(3, 4, total_events, STYLE_META_VALUE)

    ws_sum.set_row_height(4, 20)
    ws_sum.set_cell(4, 1, "Highlighted Words:", STYLE_META_LABEL)
    hl_str = ", ".join(highlight_words) if highlight_words else "(None)"
    ws_sum.set_cell(4, 2, hl_str, STYLE_META_VALUE)
    ws_sum.set_cell(4, 3, "Total Highlights:", STYLE_META_LABEL)
    total_highlights = sum(g["total_highlights"] for g in groups_data)
    ws_sum.set_cell(4, 4, total_highlights, STYLE_META_VALUE)

    ws_sum.set_row_height(5, 20)
    ws_sum.set_cell(5, 1, "AWS Region:", STYLE_META_LABEL)
    ws_sum.set_cell(5, 2, report_meta.get("region", "Default"), STYLE_META_VALUE)
    ws_sum.set_cell(5, 3, "AWS Profile:", STYLE_META_LABEL)
    ws_sum.set_cell(5, 4, report_meta.get("profile", "Default"), STYLE_META_VALUE)

    # Section title
    ws_sum.set_row_height(7, 24)
    ws_sum.set_cell(7, 1, "Log Group Overview", STYLE_SECTION_HEADER)

    # Table Header
    sum_headers = [
        ("No", STYLE_SUMMARY_HEADER_CENTER),
        ("Log Group Name", STYLE_SUMMARY_HEADER_LEFT),
        ("Worksheet Name", STYLE_SUMMARY_HEADER_LEFT),
        ("Streams", STYLE_SUMMARY_HEADER_CENTER),
        ("Total Events", STYLE_SUMMARY_HEADER_CENTER),
        ("Highlights", STYLE_SUMMARY_HEADER_CENTER),
        ("Time Window of Log Entries", STYLE_SUMMARY_HEADER_CENTER)
    ]
    ws_sum.set_row_height(8, 24)
    for c_idx, (h_title, h_style) in enumerate(sum_headers, 1):
        ws_sum.set_cell(8, c_idx, h_title, h_style)

    # Table Data
    current_sum_row = 9
    for idx, g in enumerate(groups_data, 1):
        is_zebra = (idx % 2 == 0)
        c_center = STYLE_DATA_CENTER_ZEBRA if is_zebra else STYLE_DATA_CENTER
        c_left = STYLE_DATA_LEFT_ZEBRA if is_zebra else STYLE_DATA_LEFT

        ws_sum.set_row_height(current_sum_row, 20)
        ws_sum.set_cell(current_sum_row, 1, idx, c_center)
        ws_sum.set_cell(current_sum_row, 2, g["name"], c_left)
        ws_sum.set_cell(current_sum_row, 3, g["sheet_name"], c_left)
        ws_sum.set_cell(current_sum_row, 4, g["stream_count"], c_center)
        ws_sum.set_cell(current_sum_row, 5, g["total_events"], c_center)

        # Highlight cell in red if highlights found
        if g["total_highlights"] > 0:
            ws_sum.set_cell(current_sum_row, 6, g["total_highlights"], STYLE_HIGHLIGHT_CENTER)
        else:
            ws_sum.set_cell(current_sum_row, 6, 0, c_center)

        window_str = f"{g['first_time']} ~ {g['last_time']}" if g['total_events'] > 0 else "No events"
        ws_sum.set_cell(current_sum_row, 7, window_str, c_center)
        current_sum_row += 1

    # -------------------------------------------------------------
    # 2. Log Group Sheets
    # -------------------------------------------------------------
    for g in groups_data:
        ws = group_worksheets[g["name"]]

        # Group Title Header
        ws.set_row_height(1, 28)
        ws.set_cell(1, 1, f"Log Group: {g['name']}", STYLE_TITLE)
        ws.merge_cells("A1:E1")

        # Group Info Bar
        ws.set_row_height(2, 20)
        ws.set_cell(2, 1, "Streams:", STYLE_META_LABEL)
        ws.set_cell(2, 2, f"{g['stream_count']} streams", STYLE_META_VALUE)
        ws.set_cell(2, 3, "Total Events:", STYLE_META_LABEL)
        ws.set_cell(2, 4, f"{g['total_events']} events", STYLE_META_VALUE)
        ws.set_cell(2, 5, f"Highlights: {g['total_highlights']} items", STYLE_HIGHLIGHT_CENTER if g['total_highlights'] > 0 else STYLE_META_VALUE)

        # Column Header Row (Fixed at row 4)
        group_headers = [
            ("No", STYLE_HEADER_CENTER),
            ("Timestamp (JST)", STYLE_HEADER_CENTER),
            ("Log Stream", STYLE_HEADER_LEFT),
            ("Event ID", STYLE_HEADER_CENTER),
            ("Log Message", STYLE_HEADER_LEFT)
        ]
        ws.set_row_height(4, 24)
        for c_idx, (h_title, h_style) in enumerate(group_headers, 1):
            ws.set_cell(4, c_idx, h_title, h_style)

        current_row = 5

        # Check if there are no events in this group
        if g["total_events"] == 0:
            ws.set_row_height(current_row, 24)
            ws.set_cell(current_row, 1, "-", STYLE_DATA_CENTER)
            ws.set_cell(current_row, 2, "-", STYLE_DATA_CENTER)
            ws.set_cell(current_row, 3, "-", STYLE_DATA_LEFT)
            ws.set_cell(current_row, 4, "-", STYLE_DATA_CENTER)
            ws.set_cell(current_row, 5, "指定された時間範囲内に出力されたログイベントはありません。", STYLE_DATA_MSG)
            continue

        # Render stream blocks
        overall_event_no = 1
        for s_idx, (stream_name, events) in enumerate(g["streams"].items(), 1):
            # Stream Banner Block
            ws.set_row_height(current_row, 22)
            first_t = epoch_ms_to_jst_str(events[0]["timestamp"]) if events else "N/A"
            last_t = epoch_ms_to_jst_str(events[-1]["timestamp"]) if events else "N/A"
            banner_text = f"■ ログストリーム: {stream_name}  ({len(events)} 件 | {first_t} ~ {last_t})"
            ws.set_cell(current_row, 1, banner_text, STYLE_STREAM_BANNER)
            for c_fill in range(2, 6):
                ws.set_cell(current_row, c_fill, "", STYLE_STREAM_BANNER)
            ws.merge_cells(f"A{current_row}:E{current_row}")
            current_row += 1

            # Event rows
            for ev in events:
                msg = ev.get("message", "")
                is_hl = has_highlight(msg)
                ts_str = epoch_ms_to_jst_str(ev.get("timestamp", 0))
                ev_id = ev.get("eventId", "")

                is_zebra = (overall_event_no % 2 == 0)

                if is_hl:
                    s_no = STYLE_HIGHLIGHT_CENTER
                    s_ts = STYLE_HIGHLIGHT_CENTER
                    s_st = STYLE_HIGHLIGHT_LEFT
                    s_id = STYLE_HIGHLIGHT_CENTER
                    s_msg = STYLE_HIGHLIGHT_MSG
                else:
                    s_no = STYLE_DATA_CENTER_ZEBRA if is_zebra else STYLE_DATA_CENTER
                    s_ts = STYLE_DATA_CENTER_ZEBRA if is_zebra else STYLE_DATA_CENTER
                    s_st = STYLE_DATA_LEFT_ZEBRA if is_zebra else STYLE_DATA_LEFT
                    s_id = STYLE_DATA_CENTER_ZEBRA if is_zebra else STYLE_DATA_CENTER
                    s_msg = STYLE_DATA_MSG_ZEBRA if is_zebra else STYLE_DATA_MSG

                # Estimate row height based on message line breaks
                line_count = msg.count('\n') + 1
                row_h = max(20, min(120, line_count * 18))

                ws.set_row_height(current_row, row_h)
                ws.set_cell(current_row, 1, overall_event_no, s_no)
                ws.set_cell(current_row, 2, ts_str, s_ts)
                ws.set_cell(current_row, 3, stream_name, s_st)
                ws.set_cell(current_row, 4, ev_id, s_id)
                ws.set_cell(current_row, 5, msg, s_msg)

                overall_event_no += 1
                current_row += 1

            # Spacing between stream blocks
            current_row += 1

    wb.save(output_file)

def main():
    parser = argparse.ArgumentParser(description="CloudWatch Logs to Excel Exporter (RHEL 9.8 Pure Python)")
    parser.add_argument("-b", "--start-time", dest="start_time", help="Start time (e.g. '2026-10-04 00:00:00', '1h', '24h')")
    parser.add_argument("-e", "--end-time", dest="end_time", help="End time (e.g. '2026-10-04 06:00:00')")
    parser.add_argument("-o", "--output", dest="output", default=None, help="Output Excel file path (.xlsx)")
    parser.add_argument("-H", "--highlight", dest="highlight", default="", help="Comma-separated highlight words (e.g. 'ERROR,WARN,FATAL')")
    parser.add_argument("-p", "--log-group-prefix", dest="log_group_prefix", default=None, help="Filter log groups by prefix")
    parser.add_argument("-g", "--log-groups", dest="log_groups", default="", help="Comma-separated specific log group names")
    parser.add_argument("-r", "--region", dest="region", default=None, help="AWS Region (e.g. ap-northeast-1)")
    parser.add_argument("--profile", dest="profile", default=None, help="AWS CLI profile name")
    parser.add_argument("--mock-data", dest="mock_data", default=None, help="Path to mock JSON data file (for testing/offline)")

    args = parser.parse_args()

    # Parse start and end time
    start_ms = parse_time_to_epoch_ms(args.start_time, default_now=False)
    end_ms = parse_time_to_epoch_ms(args.end_time, default_now=True)

    # Default start_ms to 1 hour before end_ms if not specified
    if start_ms is None:
        start_ms = end_ms - (3600 * 1000)

    if start_ms > end_ms:
        sys.stderr.write(f"[ERROR] Start time ({epoch_ms_to_jst_str(start_ms)}) cannot be after end time ({epoch_ms_to_jst_str(end_ms)}).\n")
        sys.exit(1)

    # Highlight words
    highlight_words = []
    if args.highlight:
        for item in args.highlight.split(','):
            cleaned = item.strip()
            if cleaned:
                highlight_words.append(cleaned)

    # Precompile highlight regex for count
    hl_patterns = [re.compile(re.escape(w), re.IGNORECASE) for w in highlight_words]
    def count_highlights(msg):
        if not msg or not hl_patterns:
            return 0
        return 1 if any(p.search(msg) for p in hl_patterns) else 0

    # Output file
    output_file = args.output
    if not output_file:
        now_str = datetime.datetime.now(JST).strftime("%Y%m%d_%H%M%S")
        output_file = f"cloudwatch_logs_{now_str}.xlsx"

    groups_data = []

    # Check for mock data
    if args.mock_data:
        print(f"[INFO] Using mock data from: {args.mock_data}")
        with open(args.mock_data, "r", encoding="utf-8") as f:
            mock_json = json.load(f)

        for g_item in mock_json.get("logGroups", []):
            g_name = g_item.get("name")
            raw_streams = g_item.get("streams", {})
            total_evs = sum(len(evs) for evs in raw_streams.values())
            hl_evs = sum(count_highlights(ev.get("message", "")) for evs in raw_streams.values() for ev in evs)

            all_ts = [ev["timestamp"] for evs in raw_streams.values() for ev in evs if "timestamp" in ev]
            first_t = epoch_ms_to_jst_str(min(all_ts)) if all_ts else "-"
            last_t = epoch_ms_to_jst_str(max(all_ts)) if all_ts else "-"

            groups_data.append({
                "name": g_name,
                "sheet_name": g_name,
                "streams": raw_streams,
                "stream_count": len(raw_streams),
                "total_events": total_evs,
                "total_highlights": hl_evs,
                "first_time": first_t,
                "last_time": last_t
            })
    else:
        # Live AWS query
        specific_groups = [g.strip() for g in args.log_groups.split(',')] if args.log_groups else None
        print(f"[INFO] Fetching log groups from AWS CloudWatch Logs...")
        try:
            log_group_names = get_log_groups(
                prefix=args.log_group_prefix,
                specific_groups=specific_groups,
                region=args.region,
                profile=args.profile
            )
        except Exception as e:
            sys.stderr.write(f"[ERROR] Failed to list log groups: {e}\n")
            sys.exit(1)

        print(f"[INFO] Found {len(log_group_names)} log group(s).")
        if not log_group_names:
            print("[WARN] No log groups matched the criteria.")

        for idx, g_name in enumerate(log_group_names, 1):
            print(f"[{idx}/{len(log_group_names)}] Fetching events for: {g_name} ...", end=" ", flush=True)
            try:
                streams_dict = get_log_events_for_group(
                    g_name,
                    start_ms=start_ms,
                    end_ms=end_ms,
                    region=args.region,
                    profile=args.profile
                )
                total_evs = sum(len(evs) for evs in streams_dict.values())
                hl_evs = sum(count_highlights(ev.get("message", "")) for evs in streams_dict.values() for ev in evs)

                all_ts = [ev["timestamp"] for evs in streams_dict.values() for ev in evs if "timestamp" in ev]
                first_t = epoch_ms_to_jst_str(min(all_ts)) if all_ts else "-"
                last_t = epoch_ms_to_jst_str(max(all_ts)) if all_ts else "-"

                groups_data.append({
                    "name": g_name,
                    "sheet_name": g_name,
                    "streams": streams_dict,
                    "stream_count": len(streams_dict),
                    "total_events": total_evs,
                    "total_highlights": hl_evs,
                    "first_time": first_t,
                    "last_time": last_t
                })
                print(f"Done ({len(streams_dict)} streams, {total_evs} events, {hl_evs} highlights)")
            except Exception as e:
                print(f"Failed!\n[WARN] Could not fetch events for {g_name}: {e}")
                # Still record group with 0 events
                groups_data.append({
                    "name": g_name,
                    "sheet_name": g_name,
                    "streams": {},
                    "stream_count": 0,
                    "total_events": 0,
                    "total_highlights": 0,
                    "first_time": "-",
                    "last_time": "-"
                })

    report_meta = {
        "generated_at": datetime.datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST"),
        "start_time_str": epoch_ms_to_jst_str(start_ms),
        "end_time_str": epoch_ms_to_jst_str(end_ms),
        "region": args.region or "Default (AWS CLI)",
        "profile": args.profile or "Default"
    }

    print(f"\n[INFO] Generating styled Excel report: {output_file} ...")
    build_excel_report(output_file, report_meta, groups_data, highlight_words)
    print(f"[SUCCESS] Excel report successfully generated: {os.path.abspath(output_file)}")

if __name__ == "__main__":
    main()
