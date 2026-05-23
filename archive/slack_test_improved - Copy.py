import warnings
warnings.filterwarnings("ignore", category=RuntimeWarning)
warnings.filterwarnings("ignore", message=".*MINGW-W64.*")

import json
import re
import html
import argparse
import pandas as pd
from datetime import datetime, timedelta
from slack_sdk import WebClient
from openai import OpenAI
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, PageBreak

# -------------------------
# Load configuration
# -------------------------

with open("config.json") as f:
    config = json.load(f)

SLACK_TOKEN = config["slack_token"]
CHANNEL_ID = config["channel_id"]
OPENAI_API_KEY = config["openai_api_key"]

client = WebClient(token=SLACK_TOKEN)
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# -------------------------
# Load Slack IDs from Excel
# -------------------------

def load_slack_ids_from_excel():
    """Load Slack member IDs from Excel file"""
    try:
        df = pd.read_excel("SlackIds.xlsx")
        id_mapping = {}
        for _, row in df.iterrows():
            if pd.notna(row['Name']) and pd.notna(row['Slack members id']):
                id_mapping[row['Slack members id']] = row['Name']
        print(f"Loaded {len(id_mapping)} user mappings from Excel")
        return id_mapping
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return {}

EXCEL_ID_MAPPING = load_slack_ids_from_excel()

# -------------------------
# Parse date argument
# -------------------------

parser = argparse.ArgumentParser(description="Export and summarise Slack messages for a given date.")
parser.add_argument(
    "date",
    nargs="?",
    default=None,
    help="Date to fetch messages for, in DD-MM-YYYY format (defaults to today)"
)
args = parser.parse_args()

if args.date:
    try:
        target_date = datetime.strptime(args.date, "%d-%m-%Y")
    except ValueError:
        raise ValueError(f"Invalid date '{args.date}'. Please use DD-MM-YYYY format (e.g. 15-03-2026).")
else:
    target_date = datetime.now()

# -------------------------
# Time range (full day)
# -------------------------

now = target_date
start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
end_of_day = start_of_day + timedelta(days=1)

oldest_ts = start_of_day.timestamp()
latest_ts = end_of_day.timestamp()

# -------------------------
# Username cache
# -------------------------

user_cache = {}

def get_username(user_id):
    """Get username from cache, Slack API, or Excel mapping"""
    if user_id in user_cache:
        return user_cache[user_id]

    if user_id in EXCEL_ID_MAPPING:
        name = EXCEL_ID_MAPPING[user_id]
        user_cache[user_id] = name
        return name

    try:
        user = client.users_info(user=user_id)
        name = user["user"]["real_name"]
        user_cache[user_id] = name
        return name
    except:
        name = user_id
        user_cache[user_id] = name
        return name

# -------------------------
# Clean Slack formatting
# -------------------------

def clean_text(text):
    """Clean Slack text formatting and replace mentions"""
    if not text:
        return ""

    text = html.unescape(text)

    matches = re.findall(r"<@([A-Z0-9]+)>", text)
    for uid in matches:
        username = get_username(uid)
        text = text.replace(f"<@{uid}>", f"@{username}")

    text = text.replace("<!channel>", "@channel")
    text = text.replace("<!here>", "@here")
    text = text.replace("<!everyone>", "@everyone")

    url_matches = re.findall(r"<(http[^|>]+)(?:\|[^>]+)?>", text)
    for url in url_matches:
        text = text.replace(f"<{url}>", url)

    return text

# -------------------------
# Collect messages WITH full thread pagination
# -------------------------

print("Collecting messages from Slack...")

# threads_map: { parent_ts -> {"parent": msg, "replies": [reply, ...]} }
threads_map = {}
# standalone messages (no thread)
standalone_messages = []

cursor = None

while True:
    response = client.conversations_history(
        channel=CHANNEL_ID,
        oldest=oldest_ts,
        latest=latest_ts,
        limit=200,
        cursor=cursor
    )

    messages = response["messages"]

    for message in messages:
        ts = message["ts"]

        if message.get("reply_count", 0) > 0:
            # This message has a thread — fetch ALL replies with pagination
            thread_replies = []
            thread_cursor = None

            while True:
                try:
                    replies_response = client.conversations_replies(
                        channel=CHANNEL_ID,
                        ts=ts,
                        limit=200,
                        cursor=thread_cursor
                    )
                    # replies_response["messages"][0] is always the parent — skip it
                    thread_replies.extend(replies_response["messages"][1:])
                    thread_cursor = replies_response.get("response_metadata", {}).get("next_cursor")
                    if not thread_cursor:
                        break
                except Exception as e:
                    print(f"Error fetching replies for message {ts}: {e}")
                    break

            threads_map[ts] = {
                "parent": message,
                "replies": thread_replies
            }
        else:
            # Check if this is a reply that belongs to a thread already captured
            # (conversations_history sometimes returns reply messages too)
            if message.get("thread_ts") and message.get("thread_ts") != ts:
                pass  # Will be captured via conversations_replies above
            else:
                standalone_messages.append(message)

    cursor = response.get("response_metadata", {}).get("next_cursor")
    if not cursor:
        break

total_msgs = len(standalone_messages) + sum(
    1 + len(t["replies"]) for t in threads_map.values()
)
print(f"Collected {total_msgs} messages ({len(threads_map)} threads, {len(standalone_messages)} standalone)")

# -------------------------
# Fetch old threads with new activity on target date
# -------------------------

lookback_days = config.get("lookback_days", 30)
lookback_oldest_ts = (start_of_day - timedelta(days=lookback_days)).timestamp()

print(f"Scanning last {lookback_days} days for old threads updated on {now.strftime('%Y-%m-%d')}...")

old_threads_updated = []  # list of {"parent": msg, "replies": [...], "permalink": str}

old_cursor = None
while True:
    old_response = client.conversations_history(
        channel=CHANNEL_ID,
        oldest=lookback_oldest_ts,
        latest=oldest_ts,        # only messages BEFORE the target date
        limit=200,
        cursor=old_cursor
    )

    for message in old_response["messages"]:
        if message.get("reply_count", 0) > 0:
            latest_reply_ts = float(message.get("latest_reply", 0))
            if oldest_ts <= latest_reply_ts < latest_ts:
                # This old thread received replies on the target date
                ts = message["ts"]
                thread_replies = []
                thread_cursor = None
                while True:
                    try:
                        replies_response = client.conversations_replies(
                            channel=CHANNEL_ID,
                            ts=ts,
                            limit=200,
                            cursor=thread_cursor
                        )
                        thread_replies.extend(replies_response["messages"][1:])
                        thread_cursor = replies_response.get("response_metadata", {}).get("next_cursor")
                        if not thread_cursor:
                            break
                    except Exception as e:
                        print(f"Error fetching replies for old thread {ts}: {e}")
                        break

                permalink = ""
                try:
                    pl_response = client.chat_getPermalink(channel=CHANNEL_ID, message_ts=ts)
                    permalink = pl_response.get("permalink", "")
                except Exception as e:
                    print(f"Could not get permalink for {ts}: {e}")

                old_threads_updated.append({
                    "parent": message,
                    "replies": thread_replies,
                    "permalink": permalink
                })

    old_cursor = old_response.get("response_metadata", {}).get("next_cursor")
    if not old_cursor:
        break

print(f"Found {len(old_threads_updated)} old thread(s) with activity on {now.strftime('%Y-%m-%d')}")

# -------------------------
# Format a single message line
# -------------------------

def format_message_line(message, indent=""):
    ts = float(message["ts"])
    dt = datetime.fromtimestamp(ts).strftime("%H:%M")
    text = clean_text(message.get("text", ""))
    user_id = message.get("user")
    username = get_username(user_id) if user_id else "SYSTEM"

    # Capture any file attachments
    files = message.get("files", [])
    file_info = ""
    if files:
        file_names = [f.get("name", "unknown file") for f in files]
        file_info = f" [Attachment(s): {', '.join(file_names)}]"

    # Capture link/rich preview attachments
    attachments = message.get("attachments", [])
    attachment_info = ""
    if attachments:
        titles = [a.get("title") or a.get("fallback", "") for a in attachments if a.get("title") or a.get("fallback")]
        if titles:
            attachment_info = f" [Preview: {'; '.join(titles)}]"

    return f"{indent}[{dt}] {username}: {text}{file_info}{attachment_info}"

# -------------------------
# Build structured raw output (threads preserved)
# -------------------------

# Collect all top-level items (parent messages + standalone) sorted by timestamp
all_top_level = []

for ts, thread in threads_map.items():
    all_top_level.append(("thread", float(ts), thread))

for msg in standalone_messages:
    all_top_level.append(("standalone", float(msg["ts"]), msg))

all_top_level.sort(key=lambda x: x[1])

# Build raw output lines
raw_output = []
thread_counter = 0

for item_type, _, data in all_top_level:
    if item_type == "standalone":
        raw_output.append(format_message_line(data))
    else:
        thread_counter += 1
        parent = data["parent"]
        replies = data["replies"]
        reply_count = len(replies)

        raw_output.append("")
        raw_output.append(f"┌── THREAD #{thread_counter} ({reply_count} repl{'y' if reply_count == 1 else 'ies'}) ──────────────────────────────────────")
        raw_output.append(format_message_line(parent, indent="│ "))

        if replies:
            raw_output.append("│   └─ Replies:")
            for reply in replies:
                raw_output.append(format_message_line(reply, indent="│      "))

        raw_output.append("└─────────────────────────────────────────────────────────────────────────")
        raw_output.append("")

raw_text = "\n".join(raw_output)

# -------------------------
# Build old-threads section (deterministic, not AI-generated)
# -------------------------

def build_old_threads_section():
    lines = []
    lines.append("=" * 80)
    lines.append(f"OLD THREADS WITH NEW ACTIVITY ON {now.strftime('%Y-%m-%d')} ({len(old_threads_updated)} found, lookback: {lookback_days} days)")
    lines.append("=" * 80)

    if not old_threads_updated:
        lines.append("\nNo old threads had new activity on this date.")
        return "\n".join(lines)

    for i, thread in enumerate(old_threads_updated, 1):
        parent = thread["parent"]
        replies = thread["replies"]
        permalink = thread["permalink"]

        parent_ts = float(parent["ts"])
        parent_dt = datetime.fromtimestamp(parent_ts).strftime("%Y-%m-%d %H:%M")
        parent_user = get_username(parent.get("user")) if parent.get("user") else "SYSTEM"
        parent_text = clean_text(parent.get("text", ""))

        # Separate earlier replies from new replies on target date
        earlier_replies = [r for r in replies if float(r["ts"]) < oldest_ts]
        new_replies = [r for r in replies if oldest_ts <= float(r["ts"]) < latest_ts]

        lines.append("")
        lines.append(f"┌── OLD THREAD #{i}  (originally posted {parent_dt}) ──────────────────────────────")
        lines.append(f"│  Author  : {parent_user}")
        lines.append(f"│  Message : {parent_text}")
        if permalink:
            lines.append(f"│  Link    : {permalink}")
        if earlier_replies:
            lines.append(f"│  [{len(earlier_replies)} earlier repl{'y' if len(earlier_replies) == 1 else 'ies'} before {now.strftime('%Y-%m-%d')} not shown]")
        if new_replies:
            lines.append(f"│  New repl{'y' if len(new_replies) == 1 else 'ies'} on {now.strftime('%Y-%m-%d')} ({len(new_replies)}):")
            for reply in new_replies:
                lines.append(format_message_line(reply, indent="│    "))
        lines.append("└─────────────────────────────────────────────────────────────────────────")

    return "\n".join(lines)

old_threads_section = build_old_threads_section()

# -------------------------
# Also build a flat summary-friendly version that retains thread context labels
# -------------------------

def build_summary_text():
    """Build a clean, AI-friendly version of messages that preserves thread context"""
    lines = []

    for item_type, _, data in all_top_level:
        if item_type == "standalone":
            lines.append(format_message_line(data))
        else:
            parent = data["parent"]
            replies = data["replies"]

            parent_ts = float(parent["ts"])
            parent_dt = datetime.fromtimestamp(parent_ts).strftime("%H:%M")
            parent_user = get_username(parent.get("user")) if parent.get("user") else "SYSTEM"
            parent_text = clean_text(parent.get("text", ""))

            lines.append(f"\n[THREAD STARTED AT {parent_dt}]")
            lines.append(f"  Original message — {parent_user}: {parent_text}")

            if replies:
                lines.append(f"  Thread replies ({len(replies)}):")
                for reply in replies:
                    lines.append(format_message_line(reply, indent="    "))

            lines.append("[END THREAD]\n")

    return "\n".join(lines)

summary_input_text = build_summary_text()

# -------------------------
# Detect project/job codes in threads
# -------------------------

def extract_project_codes(text_block):
    """
    Scan a block of text for project/job codes.
    Returns a list of (keyword, code) tuples found.
    A match is any line containing 'project' or 'job' AND a number on the same line.
    Handles separators: :-  -  :  or plain space.
    """
    pattern = re.compile(r'(?i)\b(project|job)\b[^\n]*?(\d+)')
    found = []
    for line in text_block.split('\n'):
        match = pattern.search(line)
        if match:
            keyword = match.group(1).lower()
            code = match.group(2)
            if (keyword, code) not in found:
                found.append((keyword, code))
    return found

threads_with_code = []    # (thread_num, ts, codes, parent_snippet)
threads_without_code = [] # (thread_num, ts, parent_snippet)

_thread_num = 0
for item_type, ts, data in all_top_level:
    if item_type == "thread":
        _thread_num += 1
        parent = data["parent"]
        all_msgs = [parent] + data["replies"]
        full_text = "\n".join(clean_text(m.get("text", "")) for m in all_msgs)
        codes = extract_project_codes(full_text)
        snippet = clean_text(parent.get("text", ""))[:80]
        if codes:
            threads_with_code.append((_thread_num, ts, codes, snippet))
        else:
            threads_without_code.append((_thread_num, ts, snippet))

def build_project_code_section():
    lines = []
    lines.append("=" * 80)
    lines.append("PROJECT CODE TRACKING")
    lines.append("=" * 80)

    lines.append(f"\nTHREADS WITH PROJECT/JOB CODE ({len(threads_with_code)})")
    if threads_with_code:
        for thread_num, ts, codes, snippet in threads_with_code:
            dt = datetime.fromtimestamp(ts).strftime("%H:%M")
            code_strs = ", ".join(f"{kw.upper()} {code}" for kw, code in codes)
            lines.append(f"  Thread #{thread_num} [{dt}] — Code(s): {code_strs}")
            ellipsis = "..." if len(snippet) == 80 else ""
            lines.append(f"    \"{snippet}{ellipsis}\"")
    else:
        lines.append("  (none)")

    lines.append(f"\nTHREADS WITHOUT PROJECT/JOB CODE ({len(threads_without_code)})")
    if threads_without_code:
        for thread_num, ts, snippet in threads_without_code:
            dt = datetime.fromtimestamp(ts).strftime("%H:%M")
            ellipsis = "..." if len(snippet) == 80 else ""
            lines.append(f"  Thread #{thread_num} [{dt}] — \"{snippet}{ellipsis}\"")
    else:
        lines.append("  (none)")

    return "\n".join(lines)

project_code_section = build_project_code_section()

# -------------------------
# Save raw text to file
# -------------------------

timestamp_str = now.strftime("%Y%m%d")
raw_filename = f"slack_raw_{timestamp_str}.txt"
with open(raw_filename, "w", encoding="utf-8") as f:
    f.write(f"SLACK RAW EXPORT — {now.strftime('%Y-%m-%d')}\n")
    f.write(f"Channel: {CHANNEL_ID}\n")
    f.write(f"Total messages: {total_msgs} | Threads: {len(threads_map)} | Standalone: {len(standalone_messages)}\n")
    f.write("=" * 80 + "\n\n")
    f.write(raw_text)

print(f"Raw output saved to {raw_filename}")

# -------------------------
# Generate summary using OpenAI
# -------------------------

def generate_summary(messages_text):
    """Generate a detailed summary using OpenAI, with full thread awareness"""

    prompt = f"""You are analyzing Slack messages exported from a team channel for {now.strftime('%Y-%m-%d')}.

The messages below preserve thread structure — look for [THREAD STARTED AT ...] blocks to understand full conversations.
Standalone messages outside thread blocks are direct channel posts.

Provide a DETAILED and COMPREHENSIVE summary structured as follows:

---

## 1. ISSUES REPORTED
For each issue:
- Description of the problem
- Reported by (name) at (time)
- Current status: resolved / unresolved / in-progress (based on thread replies)
- Resolution or next step if mentioned

## 2. TASKS ASSIGNED
For each task:
- Task description
- Assigned by → Assigned to
- Deadline or urgency if mentioned
- Status (acknowledged / pending / completed)

## 3. DECISIONS MADE
- Any decisions or conclusions reached during discussions
- Include which thread or conversation the decision came from

## 4. OPEN QUESTIONS (Unanswered)
- Questions that were asked but not answered
- Who asked, and at what time
- Tag as [URGENT] if the question seems time-sensitive

## 5. KEY DISCUSSION HIGHLIGHTS
- Summarize the most important thread discussions (not just top-level messages)
- Include back-and-forth context that adds meaning
- Note any disagreements, clarifications, or important context shared in replies

## 6. THREAD-BY-THREAD SUMMARY
List every thread and every standalone message in chronological order.
For each item use this format:

### [HH:MM] — <one-line description of the original message>
**Original message:** <full or near-full text of the opening message, attributed to the sender>
**Thread summary:** <concise summary of the replies and how the discussion evolved — who said what, what was resolved, what is still open. If there are no replies, write "No replies.">

---

Be thorough. Do not skip information from thread replies — they often contain the most important details.
Mention names wherever possible for accountability.

MESSAGES:
{messages_text}
"""

    try:
        response = openai_client.chat.completions.create(
            model="gpt-3.5-turbo-16k",  # Use 16k model to handle larger thread content
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a meticulous assistant that produces detailed, structured summaries "
                        "of Slack conversations. You never skip thread replies and always attribute "
                        "statements to the correct person. You flag unresolved issues clearly."
                    )
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=4000  # Section 6 (thread-by-thread) adds significant length
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error generating summary: {e}")
        return "Error generating summary. Please check OpenAI API key and try again."

print("Generating detailed summary using OpenAI...")
summary = generate_summary(summary_input_text)

# -------------------------
# Save processed summary to file
# -------------------------

processed_filename = f"slack_processed_{timestamp_str}.txt"
with open(processed_filename, "w", encoding="utf-8") as f:
    f.write("=" * 80 + "\n")
    f.write("SLACK MESSAGES SUMMARY\n")
    f.write(f"Date: {now.strftime('%Y-%m-%d')}\n")
    f.write(f"Total messages: {total_msgs} | Threads: {len(threads_map)} | Standalone: {len(standalone_messages)}\n")
    f.write("=" * 80 + "\n\n")
    f.write(summary)
    f.write("\n\n")
    f.write(project_code_section)
    f.write("\n\n")
    f.write(old_threads_section)
    f.write("\n\n" + "=" * 80 + "\n")
    f.write("RAW MESSAGES (Thread-Structured)\n")
    f.write("=" * 80 + "\n\n")
    f.write(raw_text)

print(f"Processed summary saved to {processed_filename}")

# -------------------------
# Generate PDF with clickable index
# -------------------------

def safe(text):
    """Escape special XML characters for use inside ReportLab Paragraph markup."""
    return html.escape(str(text)) if text else ""

def generate_pdf(pdf_filename):
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm,
        title=f"Slack Summary {now.strftime('%Y-%m-%d')}",
    )

    base_styles = getSampleStyleSheet()

    title_style  = ParagraphStyle("PTitle",  parent=base_styles["Title"],
                                  fontSize=20, spaceAfter=6)
    date_style   = ParagraphStyle("PDate",   parent=base_styles["Normal"],
                                  fontSize=11, textColor=colors.grey, spaceAfter=14)
    h1_style     = ParagraphStyle("PH1",     parent=base_styles["Heading1"],
                                  fontSize=14, spaceBefore=18, spaceAfter=6,
                                  textColor=colors.HexColor("#1a1a6e"))
    h2_style     = ParagraphStyle("PH2",     parent=base_styles["Heading2"],
                                  fontSize=11, spaceBefore=10, spaceAfter=4,
                                  textColor=colors.HexColor("#333366"))
    toc_style    = ParagraphStyle("PTOC",    parent=base_styles["Normal"],
                                  fontSize=11, spaceAfter=7, leftIndent=10,
                                  textColor=colors.HexColor("#0000cc"))
    body_style   = ParagraphStyle("PBody",   parent=base_styles["Normal"],
                                  fontSize=9,  spaceAfter=3, leading=14)
    bullet_style = ParagraphStyle("PBullet", parent=base_styles["Normal"],
                                  fontSize=9,  spaceAfter=3, leading=14, leftIndent=14)
    mono_style   = ParagraphStyle("PMono",   parent=base_styles["Normal"],
                                  fontName="Courier", fontSize=7.5, spaceAfter=1,
                                  leading=11)

    story = []

    # ── Title block ──────────────────────────────────────────────────────────
    story.append(Paragraph("Slack Channel Summary", title_style))
    story.append(Paragraph(f"Date: {now.strftime('%Y-%m-%d')} &nbsp;|&nbsp; "
                            f"Threads: {len(threads_map)} &nbsp;|&nbsp; "
                            f"Total messages: {total_msgs}", date_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cccccc")))
    story.append(Spacer(1, 0.25*cm))

    # ── Table of Contents ────────────────────────────────────────────────────
    story.append(Paragraph('<a name="toc"/>Table of Contents', h1_style))
    toc_entries = [
        ("1.  AI-Generated Summary",             "sec_summary"),
        ("2.  Project Code Tracking",            "sec_project"),
        ("3.  Old Threads with New Activity",    "sec_old"),
        ("4.  Raw Messages (Thread-Structured)", "sec_raw"),
    ]
    for label, anchor in toc_entries:
        story.append(Paragraph(f'<a href="#{anchor}" color="#0000cc">{safe(label)}</a>', toc_style))

    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cccccc")))

    # ── Helper: render the AI summary text with basic markdown-like headings ─
    h3_style = ParagraphStyle("PH3", parent=base_styles["Heading3"],
                               fontSize=10, spaceBefore=10, spaceAfter=3,
                               textColor=colors.HexColor("#555500"))
    bold_label_style = ParagraphStyle("PBoldLabel", parent=base_styles["Normal"],
                                      fontSize=9, spaceAfter=3, leading=14)

    def render_summary_lines(text):
        for raw_line in text.split("\n"):
            line = raw_line.rstrip()
            if not line:
                story.append(Spacer(1, 4))
            elif line.startswith("### "):
                story.append(Paragraph(safe(line[4:]), h3_style))
            elif line.startswith("## "):
                story.append(Paragraph(safe(line[3:]), h2_style))
            elif line.startswith("# "):
                story.append(Paragraph(safe(line[2:]), h1_style))
            elif line.startswith("**") and line.endswith("**"):
                story.append(Paragraph(f"<b>{safe(line[2:-2])}</b>", bold_label_style))
            elif re.match(r"^\*\*[^*]+\*\*", line):
                # Bold label followed by text e.g. **Original message:** some text
                formatted = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", safe(line))
                story.append(Paragraph(formatted, body_style))
            elif line.startswith("- ") or line.startswith("* "):
                story.append(Paragraph(f"&#8226; {safe(line[2:])}", bullet_style))
            else:
                story.append(Paragraph(safe(line), body_style))

    # ── Helper: render pre-formatted mono lines, skipping pure separator lines ─
    def render_mono_lines(text):
        for raw_line in text.split("\n"):
            line = raw_line.rstrip()
            if set(line) <= {"=", "-", "─", "│", "┌", "└", " ", ""}:
                # Keep structural lines but render them smaller
                story.append(Paragraph(safe(line) or " ", mono_style))
            else:
                story.append(Paragraph(safe(line) or " ", mono_style))

    # ── Section 1: AI Summary ─────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph('<a name="sec_summary"/>1. AI-Generated Summary', h1_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa")))
    story.append(Spacer(1, 0.2*cm))
    render_summary_lines(summary)

    # ── Section 2: Project Code Tracking ─────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph('<a name="sec_project"/>2. Project Code Tracking', h1_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa")))
    story.append(Spacer(1, 0.2*cm))
    render_mono_lines(project_code_section)

    # ── Section 3: Old Threads with New Activity ──────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph('<a name="sec_old"/>3. Old Threads with New Activity', h1_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa")))
    story.append(Spacer(1, 0.2*cm))
    # Render permalink lines as actual clickable hyperlinks
    for raw_line in old_threads_section.split("\n"):
        line = raw_line.rstrip()
        if line.startswith("│  Link    :"):
            url = line.split(":", 1)[1].strip()
            story.append(Paragraph(
                f'│  Link    : <a href="{safe(url)}" color="#0000cc">{safe(url)}</a>',
                mono_style))
        else:
            story.append(Paragraph(safe(line) or " ", mono_style))

    # ── Section 4: Raw Messages ───────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph('<a name="sec_raw"/>4. Raw Messages (Thread-Structured)', h1_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa")))
    story.append(Spacer(1, 0.2*cm))
    render_mono_lines(raw_text)

    doc.build(story)

pdf_filename = f"slack_summary_{timestamp_str}.pdf"
try:
    generate_pdf(pdf_filename)
    print(f"PDF saved to {pdf_filename}")
except Exception as e:
    print(f"PDF generation failed: {e}")
    pdf_filename = None

# -------------------------
# Display summary in console
# -------------------------

print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print(summary)
print("\n")
print(project_code_section)
print("\n")
print(old_threads_section)
print("\n" + "=" * 80)
print(f"Files created:")
print(f"  - {raw_filename}  (structured raw export)")
print(f"  - {processed_filename}  (summary + full raw)")
if pdf_filename:
    print(f"  - {pdf_filename}  (PDF with clickable index)")
