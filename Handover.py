from __future__ import annotations

import datetime as dt
import io
import sqlite3
from pathlib import Path
from urllib.parse import quote
from typing import Dict, Iterable, Optional

import pandas as pd
import streamlit as st

try:  # Plotly is optional; degrade gracefully if unavailable
    import plotly.express as px  # type: ignore
except ImportError:  # pragma: no cover - visual enhancement only
    px = None  # type: ignore[assignment]

try:  # Prefer xlsxwriter for richer formatting, otherwise fall back
    import xlsxwriter  # type: ignore  # noqa: F401

    EXCEL_ENGINE = "xlsxwriter"
except ImportError:  # pragma: no cover - platform dependent
    EXCEL_ENGINE = "openpyxl"


DB_PATH = Path(__file__).with_name("handover_tracker.db")
STATUS_OPTIONS = ["Not Started", "In Progress", "Waiting", "Completed", "Archived"]
STATUS_BADGES = {
    "Not Started": "🟥 Not Started",
    "In Progress": "🟨 In Progress",
    "Waiting": "🟦 Waiting",
    "Completed": "🟩 Completed",
    "Archived": "⬜ Archived",
}

STATUS_COLORS = {
    "Not Started": "#d62728",
    "In Progress": "#ffbf00",
    "Waiting": "#1f77b4",
    "Completed": "#2ca02c",
    "Archived": "#7f7f7f",
}


def status_badge(status: str) -> str:
    return STATUS_BADGES.get(status, status)


def status_color(status: str) -> str:
    return STATUS_COLORS.get(status, "#636EFA")


def build_excel_report(df: pd.DataFrame) -> bytes:
    """Generate an Excel workbook with handover data."""

    export_df = df.copy()
    export_df["Meeting Date"] = export_df["meeting_date"].apply(lambda d: d.strftime("%d %b %Y"))
    export_df["Status"] = export_df["status"].apply(status_badge)
    export_df["Created"] = export_df["created_at"].dt.strftime("%d %b %Y %H:%M")
    export_df["Last Updated"] = export_df["updated_at"].dt.strftime("%d %b %Y %H:%M")

    columns = [
        "person",
        "topic",
        "Meeting Date",
        "details",
        "Status",
        "next_steps",
        "Created",
        "Last Updated",
    ]

    export_df = export_df[columns]
    export_df = export_df.rename(
        columns={
            "person": "Owner",
            "topic": "Topic",
            "details": "Key Details",
            "next_steps": "Next Steps",
        }
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
        export_df.to_excel(writer, index=False, sheet_name="Handover Register")
        if EXCEL_ENGINE == "xlsxwriter":
            workbook = writer.book
            worksheet = writer.sheets["Handover Register"]

            header_format = workbook.add_format(
                {"bold": True, "bg_color": "#003366", "font_color": "#FFFFFF", "align": "center"}
            )
            for col_num, value in enumerate(export_df.columns):
                worksheet.write(0, col_num, value, header_format)
                column_width = max(15, int(export_df[value].astype(str).str.len().max()) + 2)
                worksheet.set_column(col_num, col_num, min(column_width, 60))

            date_format = workbook.add_format({"num_format": "dd mmm yyyy"})
            worksheet.set_column("C:C", 18, date_format)

    buffer.seek(0)
    return buffer.read()


def build_summary(df: pd.DataFrame) -> str:
    total = len(df)
    completed = int((df["status"] == "Completed").sum())
    waiting = int((df["status"] == "Waiting").sum())
    in_progress = int((df["status"] == "In Progress").sum())

    upcoming_window = dt.date.today() + dt.timedelta(days=7)
    upcoming = int((df["meeting_date"] <= upcoming_window).sum())

    summary_lines = [
        "# Handover summary",
        f"- Total topics: {total}",
        f"- Completed: {completed}",
        f"- In progress: {in_progress}",
        f"- Waiting on others: {waiting}",
        f"- Due within 7 days: {upcoming}",
        "",
        "## Highlights",
    ]

    top_items = (
        df.sort_values("meeting_date")
        .head(5)
        .assign(
            badge=lambda d: d["status"].apply(status_badge),
            due=lambda d: d["meeting_date"].apply(lambda date: date.strftime("%d %b")),
        )
    )

    if top_items.empty:
        summary_lines.append("No topics recorded yet.")
    else:
        for row in top_items.itertuples():
            summary_lines.append(
                f"- {row.due}: {row.topic} ({row.person}) · {row.badge}\n  Next: {row.next_steps or 'No next steps logged.'}"
            )

    summary_lines.append("")
    summary_lines.append("Generated via the Handover Tracker")
    return "\n".join(summary_lines)


def get_connection() -> sqlite3.Connection:
    """Return a SQLite connection, creating the database if needed."""

    DB_PATH.touch(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    """Ensure the handover table exists."""

    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS handover_topics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person TEXT NOT NULL,
                topic TEXT NOT NULL,
                meeting_date TEXT NOT NULL,
                details TEXT,
                status TEXT NOT NULL,
                next_steps TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )


def add_topic(
    person: str,
    topic: str,
    meeting_date: dt.date,
    details: str,
    status: str,
    next_steps: str,
) -> None:
    now_iso = dt.datetime.utcnow().isoformat(timespec="seconds")
    payload = (
        person.strip(),
        topic.strip(),
        meeting_date.isoformat(),
        details.strip(),
        status,
        next_steps.strip(),
        now_iso,
        now_iso,
    )

    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO handover_topics (
                person, topic, meeting_date, details, status, next_steps, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            payload,
        )


def update_topic(
    topic_id: int,
    updates: Dict[str, str],
    meeting_date: Optional[dt.date] = None,
) -> None:
    set_parts: list[str] = []
    bindings: list[str] = []

    for column, value in updates.items():
        set_parts.append(f"{column} = ?")
        bindings.append(value)

    if meeting_date is not None:
        set_parts.append("meeting_date = ?")
        bindings.append(meeting_date.isoformat())

    set_parts.append("updated_at = ?")
    bindings.append(dt.datetime.utcnow().isoformat(timespec="seconds"))
    bindings.append(topic_id)

    with get_connection() as conn:
        conn.execute(
            f"UPDATE handover_topics SET {' ,'.join(set_parts)} WHERE id = ?",
            bindings,
        )


def fetch_topics(filters: Optional[Dict[str, Iterable[str]]] = None) -> pd.DataFrame:
    query = "SELECT * FROM handover_topics"
    conditions: list[str] = []
    params: list[str] = []

    if filters:
        if persons := list(filters.get("person", [])):
            placeholders = ",".join("?" for _ in persons)
            conditions.append(f"person IN ({placeholders})")
            params.extend(persons)

        if statuses := list(filters.get("status", [])):
            placeholders = ",".join("?" for _ in statuses)
            conditions.append(f"status IN ({placeholders})")
            params.extend(statuses)

        if conditions:
            query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY meeting_date ASC"

    with get_connection() as conn:
        df = pd.read_sql_query(query, conn, params=params)

    if not df.empty:
        df["meeting_date"] = pd.to_datetime(df["meeting_date"]).dt.date
        df["created_at"] = pd.to_datetime(df["created_at"])
        df["updated_at"] = pd.to_datetime(df["updated_at"])

    return df


def fetch_topic(topic_id: int) -> Optional[pd.Series]:
    df = fetch_topics()
    if df.empty:
        return None
    match = df.loc[df["id"] == topic_id]
    if match.empty:
        return None
    return match.iloc[0]


@st.cache_data(show_spinner=False)
def load_topics(filters: Dict[str, Iterable[str]]) -> pd.DataFrame:
    return fetch_topics(filters)


def clear_cached_topics() -> None:
    load_topics.clear()  # type: ignore[attr-defined]


def render_add_topic() -> None:
    st.subheader("✍️ Log a new handover discussion")

    with st.form("add_topic_form", clear_on_submit=True):
        person = st.text_input("Person responsible", placeholder="e.g., Ahmed Hassan")
        topic = st.text_input("Topic", placeholder="e.g., Monthly demand planning handoff")
        meeting_date = st.date_input(
            "Target handover date", value=dt.date.today(), format="DD/MM/YYYY"
        )
        details = st.text_area(
            "Key details",
            placeholder="Summary, open questions, systems used, pending actions...",
        )
        status = st.selectbox(
            "Current status",
            STATUS_OPTIONS,
            index=0,
            format_func=status_badge,
        )
        next_steps = st.text_area(
            "Next steps / notes",
            placeholder="Follow-up tasks, documents to share, risks, etc.",
        )

        submitted = st.form_submit_button("Save topic", type="primary")
        if submitted:
            if not person.strip() or not topic.strip():
                st.error("Person and topic are required.")
            else:
                add_topic(person, topic, meeting_date, details, status, next_steps)
                clear_cached_topics()
                st.success("Topic saved and ready for review.")


def render_metrics(df: pd.DataFrame) -> None:
    total = len(df)
    in_progress = int((df["status"] == "In Progress").sum())
    blocked = int((df["status"] == "Waiting").sum())
    completed = int((df["status"] == "Completed").sum())

    upcoming_window = dt.date.today() + dt.timedelta(days=7)
    upcoming = int((df["meeting_date"] <= upcoming_window).sum())

    st.markdown("### Snapshot")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🗂️ Total topics", total)
    col2.metric("🚧 In progress", in_progress)
    col3.metric("⏳ Waiting on others", blocked)
    col4.metric("⏰ Due ≤ 7 days", upcoming)

    completion_rate = completed / total if total else 0.0
    st.progress(
        completion_rate,
        text=f"Completion rate: {completed}/{total} done" if total else "Completion rate",
    )


def render_review() -> None:
    st.subheader("📋 Review & update handover topics")

    st.sidebar.markdown("### 💡 Quick tips")
    st.sidebar.success(
        "**Focus on clarity:** add document links, escalation paths, and who needs to be informed."
    )
    st.sidebar.info(
        "**Weekly rhythm:** update status, highlight blockers, and share the report link."
    )

    with st.expander("Filters", expanded=True):
        all_df = fetch_topics()
        persons = sorted(all_df["person"].unique()) if not all_df.empty else []
        statuses = sorted(all_df["status"].unique()) if not all_df.empty else STATUS_OPTIONS

        selected_persons = st.multiselect("People", options=persons)
        selected_statuses = st.multiselect(
            "Status",
            options=statuses,
            default=[],
            format_func=status_badge,
        )

    df = load_topics({"person": selected_persons, "status": selected_statuses})

    if df.empty:
        st.info("No topics logged yet. Use the form above to add your first entry.")
        return

    render_metrics(df)

    if px is None:
        st.warning(
            "Plotly is not installed, so charts are hidden. Install `plotly` to see the visuals.",
            icon="⚠️",
        )
    else:
        status_summary = (
            df.groupby("status")["id"].count().reset_index().rename(columns={"id": "count"})
        )
        status_summary["display"] = status_summary["status"].apply(status_badge)

        pie_col, chart_col = st.columns([1, 2])
        with pie_col:
            fig = px.pie(
                status_summary,
                names="display",
                values="count",
                color="status",
                color_discrete_map=STATUS_COLORS,
                title="Status distribution",
            )
            st.plotly_chart(fig, use_container_width=True)

        with chart_col:
            timeline = (
                df.groupby("meeting_date")["id"].count().reset_index().rename(columns={"id": "Topics"})
            )
            fig_timeline = px.area(
                timeline,
                x="meeting_date",
                y="Topics",
                title="Upcoming handover timeline",
                color_discrete_sequence=["#636EFA"],
            )
            fig_timeline.update_traces(mode="lines+markers", line_shape="hv")
            st.plotly_chart(fig_timeline, use_container_width=True)

    st.markdown("### 📁 Handover register")
    st.dataframe(
        df.assign(
            meeting_date=df["meeting_date"].apply(lambda d: d.strftime("%d %b %Y")),
            created_at=df["created_at"].dt.strftime("%d %b %Y %H:%M"),
            updated_at=df["updated_at"].dt.strftime("%d %b %Y %H:%M"),
        ),
        use_container_width=True,
    )

    csv = df.to_csv(index=False)
    st.download_button(
        label="Download register (CSV)",
        data=csv,
        file_name="handover_topics.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.markdown("### 📤 Share & export")
    col_csv, col_excel, col_summary = st.columns(3)

    excel_bytes = build_excel_report(df)
    summary_markdown = build_summary(df)

    with col_csv:
        st.download_button(
            label="⬇️ CSV",
            data=csv,
            file_name="handover_topics.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_excel:
        st.download_button(
            label="📊 Excel report",
            data=excel_bytes,
            file_name="handover_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_summary:
        st.download_button(
            label="📝 Markdown summary",
            data=summary_markdown,
            file_name="handover_summary.md",
            mime="text/markdown",
            use_container_width=True,
        )

    share_subject = quote("Handover status update")
    share_body = quote(summary_markdown)
    st.markdown(
        f"[📧 Share summary via email](mailto:?subject={share_subject}&body={share_body})",
        help="Opens your email client with the summary pre-filled.",
    )

    st.markdown("### Update an entry")
    topic_options = {
        f"#{row.id} · {row.topic} ({row.person})": row.id for row in df.itertuples()
    }

    if not topic_options:
        st.info("No topics to update.")
        return

    selection = st.selectbox("Select a topic", options=list(topic_options.keys()))
    topic_id = topic_options[selection]
    selected_topic = fetch_topic(topic_id)

    if selected_topic is None:
        st.warning("Topic not found; it may have been removed.")
        return

    with st.form("update_topic_form"):
        new_person = st.text_input("Person responsible", value=selected_topic.person)
        new_topic = st.text_input("Topic", value=selected_topic.topic)
        new_meeting_date = st.date_input(
            "Target handover date",
            value=selected_topic.meeting_date,
            format="DD/MM/YYYY",
        )
        new_status = st.selectbox(
            "Status",
            STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(selected_topic.status)
            if selected_topic.status in STATUS_OPTIONS
            else 0,
            format_func=status_badge,
        )
        new_details = st.text_area("Key details", value=selected_topic.details)
        new_next_steps = st.text_area("Next steps / notes", value=selected_topic.next_steps)

        submitted = st.form_submit_button("Apply changes", type="primary")

    if submitted:
        updates = {
            "person": new_person.strip(),
            "topic": new_topic.strip(),
            "details": new_details.strip(),
            "status": new_status,
            "next_steps": new_next_steps.strip(),
        }
        update_topic(topic_id, updates, meeting_date=new_meeting_date)
        clear_cached_topics()
        st.success("Topic updated.")


def main() -> None:
    st.set_page_config(page_title="Handover Tracker", layout="wide")
    st.title("Handover planning dashboard")
    st.caption("Track conversations, owners, and next steps ahead of your transition.")

    init_db()

    tab_add, tab_review = st.tabs(["Add topic", "Review & update"])

    with tab_add:
        render_add_topic()

    with tab_review:
        render_review()


if __name__ == "__main__":
    main()

