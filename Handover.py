from __future__ import annotations

import datetime as dt
import io
import json
import sqlite3
import uuid
from pathlib import Path
from urllib.parse import quote
from typing import Dict, Iterable, List, Optional, Set, Tuple

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
    try:
        import openpyxl  # type: ignore  # noqa: F401

        EXCEL_ENGINE = "openpyxl"
    except ImportError:  # pragma: no cover - platform dependent
        EXCEL_ENGINE = None


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

try:
    from barfi.flow import Block
    from barfi.flow.streamlit import st_flow

    BARFI_AVAILABLE = True
except ImportError:  # pragma: no cover - optional dependency
    Block = None  # type: ignore[assignment]
    st_flow = None  # type: ignore[assignment]
    BARFI_AVAILABLE = False


def _graphviz_label(text: str) -> str:
    return text.replace("\\", "\\\\").replace("\n", "\\n").replace("\"", "\\\"")


def _wrap_title(title: str, max_chars: int = 40) -> str:
    clean = title.strip()
    if len(clean) <= max_chars:
        return clean

    tokens = clean.split()
    if not tokens:
        return clean

    lines: List[str] = []
    current: List[str] = []
    current_len = 0

    for token in tokens:
        next_len = current_len + len(token) + (1 if current else 0)
        if current and next_len > max_chars:
            lines.append(" ".join(current))
            current = [token]
            current_len = len(token)
        else:
            current.append(token)
            current_len = next_len

    if current:
        lines.append(" ".join(current))

    return "\n".join(lines)


def _estimate_node_width(text: str, min_width: float = 3.0, char_width: float = 0.22) -> float:
    lines = text.split("\n")
    longest = max((len(line) for line in lines), default=len(text))
    return max(min_width, longest * char_width)


def _barfi_base_blocks() -> List["Block"]:
    if not BARFI_AVAILABLE:
        return []

    step_block = Block(name="Step")
    step_block.add_input(name="In")
    step_block.add_output(name="Next")

    decision_block = Block(name="Decision")
    decision_block.add_input(name="In")
    decision_block.add_output(name="Yes")
    decision_block.add_output(name="No")

    end_block = Block(name="End")
    end_block.add_input(name="In")

    return [step_block, decision_block, end_block]


def render_barfi_editor(workflow_label: str, workflow_id: int) -> None:
    if not BARFI_AVAILABLE or st_flow is None:
        st.info(
            "Install `barfi[streamlit]` to unlock the drag-and-drop editor.\n"
            "Run: `pip install \"barfi[streamlit]\"`.",
            icon="ℹ️",
        )
        return

    st.caption(
        "Beta: Build or rearrange workflow steps visually. Drag blocks from the right-click menu, "
        "connect outputs to inputs, then review the generated schema below. Saving back to the register "
        "will be added in a follow-up iteration."
    )

    base_blocks = _barfi_base_blocks()
    if not base_blocks:
        st.warning("No Barfi blocks available.")
        return

    barfi_key = f"workflow_barfi_{workflow_id}"
    editor_result = st_flow(base_blocks, key=barfi_key)

    schema = getattr(editor_result, "editor_schema", None)
    if schema is None:
        st.info("Create nodes and click Execute in the editor to produce a schema.")
        return

    try:
        schema_payload = schema.dict()  # type: ignore[attr-defined]
    except AttributeError:
        try:
            schema_payload = schema.to_dict()  # type: ignore[attr-defined]
        except AttributeError:
            schema_payload = schema

    st.markdown("##### Generated Barfi schema")
    st.code(json.dumps(schema_payload, indent=2, default=str), language="json")
    st.info(
        "Schema persistence back to the workflow database is not yet wired. \n"
        "Use the above JSON as a reference or export template for now.",
        icon="🛠️",
    )


def build_graphviz_workflow(steps_df: pd.DataFrame) -> Optional[str]:
    if steps_df.empty:
        return None

    ordered = steps_df.sort_values("step_order")
    lines: List[str] = [
        "digraph Workflow {",
        "    rankdir=LR;",
        '    graph [splines=true, nodesep=0.8, ranksep=1.1];',
        '    node [shape=rectangle, style="rounded,filled", fontname="Helvetica", fontsize=22, fillcolor="#F7FAFC", color="#4A5568", fontcolor="#1A202C", fixedsize=false];',
        '    edge [fontname="Helvetica", fontsize=22, color="#4A5568"];',
    ]

    lines.append('    start [label="Start", shape=circle, style="filled", fillcolor="#C1E1C1", color="#2F855A", fontcolor="#1A202C", fontsize=22];')
    lines.append('    finish [label="Finish", shape=doublecircle, style="filled", fillcolor="#C1E1C1", color="#2F855A", fontcolor="#1A202C", fontsize=22];')

    def node_id(step_id: int) -> str:
        return f"step_{int(step_id)}"

    first_step_id: Optional[int] = None

    row_chunks: Dict[int, List[str]] = {}
    node_count = 0
    existing_nodes: Dict[int, str] = {}
    order_to_node: Dict[int, str] = {}

    for row in ordered.itertuples():
        if pd.isna(row.step_id):
            continue

        step_id = int(row.step_id)
        node_name = node_id(step_id)
        wrapped_title = _wrap_title(row.step_title.strip())
        title = f"Step {int(row.step_order)}: {wrapped_title}"
        width = _estimate_node_width(title)
        label = _graphviz_label(title)
        is_decision = bool(getattr(row, "no_step_id", None) and not pd.isna(row.no_step_id))
        shape = "diamond" if is_decision else "rectangle"
        border_color = "#4A5568"
        lines.append(
            f'    {node_name} [label="{label}", shape={shape}, color="{border_color}", penwidth=2, fontsize=22, width={width:.2f}];'
        )

        existing_nodes[step_id] = node_name
        order_to_node[int(row.step_order)] = node_name

        if first_step_id is None:
            first_step_id = step_id

        row_index = node_count // 4
        row_chunks.setdefault(row_index, []).append(node_name)
        node_count += 1

    if first_step_id is None:
        lines.append('    start -> finish;')
    else:
        lines.append(f'    start -> {node_id(first_step_id)};')

    if row_chunks:
        for row_index, nodes_in_row in sorted(row_chunks.items()):
            if not nodes_in_row:
                continue

            node_list = "; ".join(nodes_in_row)
            lines.append(f"    {{ rank=same; {node_list}; }}")

            if len(nodes_in_row) > 1:
                for left, right in zip(nodes_in_row, nodes_in_row[1:]):
                    lines.append(f"    {left} -> {right} [style=invis, weight=5];")

        lines.append("    { rank=source; start; }")
        lines.append("    { rank=sink; finish; }")

    for row in ordered.itertuples():
        if pd.isna(row.step_id):
            continue

        current = node_id(int(row.step_id))
        yes_target = getattr(row, "yes_step_id", None)
        no_target = getattr(row, "no_step_id", None)
        has_no_branch = bool(no_target is not None and not pd.isna(no_target))

        outgoing_edges: Set[Tuple[Optional[str], Optional[str]]] = set()

        def edge(target_raw: Optional[object], label_text: Optional[str]) -> None:
            target_name: Optional[str] = None
            target_key: Optional[int] = None
            if target_raw is not None and not pd.isna(target_raw):
                try:
                    target_key = int(target_raw)
                except (TypeError, ValueError):
                    target_key = None

            if target_key is not None:
                target_name = existing_nodes.get(target_key)
                if target_name is None:
                    target_name = order_to_node.get(target_key)

            signature = (target_name, label_text)
            if signature in outgoing_edges:
                return
            outgoing_edges.add(signature)

            attrs: List[str] = []
            if label_text:
                safe_label = _graphviz_label(label_text)
                attrs.append(f'label="{safe_label}"')
                if label_text.lower() == "yes":
                    attrs.append('color="#2F855A"')
                    attrs.append('fontcolor="#2F855A"')
                elif label_text.lower() == "no":
                    attrs.append('color="#C53030"')
                    attrs.append('fontcolor="#C53030"')
                    attrs.append('style="dotted"')
            if target_name is None:
                attr_part = f" [ {' ,'.join(attrs)} ]".replace(" [  ]", "") if attrs else ""
                lines.append(f'    {current} -> finish{attr_part};')
                return

            attr_part = f" [ {' ,'.join(attrs)} ]".replace(" [  ]", "") if attrs else ""
            lines.append(f'    {current} -> {target_name}{attr_part};')

        if has_no_branch:
            edge(yes_target, "Yes")
            edge(no_target, "No")
        else:
            edge(yes_target, None)

    lines.append('}')
    return "\n".join(lines)


def status_badge(status: str) -> str:
    return STATUS_BADGES.get(status, status)


def status_color(status: str) -> str:
    return STATUS_COLORS.get(status, "#636EFA")


def ensure_workflow_tables(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS workflow_processes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            owner TEXT,
            description TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS workflow_steps (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workflow_id INTEGER NOT NULL,
            step_order INTEGER NOT NULL,
            title TEXT NOT NULL,
            status TEXT NOT NULL,
            yes_step_id INTEGER,
            no_step_id INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (workflow_id) REFERENCES workflow_processes(id) ON DELETE CASCADE
        )
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_workflow_steps_workflow
        ON workflow_steps (workflow_id, step_order)
        """
    )
    _ensure_workflow_step_branch_columns(conn)
    conn.commit()


def _ensure_workflow_step_branch_columns(conn: sqlite3.Connection) -> None:
    existing_columns = {
        row[1] for row in conn.execute("PRAGMA table_info('workflow_steps')")
    }
    if "yes_step_id" not in existing_columns:
        conn.execute("ALTER TABLE workflow_steps ADD COLUMN yes_step_id INTEGER")
    if "no_step_id" not in existing_columns:
        conn.execute("ALTER TABLE workflow_steps ADD COLUMN no_step_id INTEGER")


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
        ensure_workflow_tables(conn)


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
    if EXCEL_ENGINE is None:
        raise RuntimeError("No Excel writer engine available. Install xlsxwriter or openpyxl.")

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


def add_workflow_process(name: str, owner: str, description: str, steps: List[str]) -> None:
    init_db()
    now_iso = dt.datetime.utcnow().isoformat(timespec="seconds")
    payload = (name.strip(), owner.strip(), description.strip(), now_iso)

    for attempt in (1, 2):
        try:
            with get_connection() as conn:
                ensure_workflow_tables(conn)
                cursor = conn.execute(
                    """
                    INSERT INTO workflow_processes (name, owner, description, created_at)
                    VALUES (?, ?, ?, ?)
                    """,
                    payload,
                )
                workflow_id = cursor.lastrowid

                inserted_ids: List[int] = []
                for order, title in enumerate(steps, start=1):
                    step_cursor = conn.execute(
                        """
                        INSERT INTO workflow_steps (
                            workflow_id, step_order, title, status, yes_step_id, no_step_id, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            workflow_id,
                            order,
                            title.strip(),
                            "Not Started",
                            None,
                            None,
                            now_iso,
                            now_iso,
                        ),
                    )
                    inserted_ids.append(int(step_cursor.lastrowid))

                for idx, step_id in enumerate(inserted_ids):
                    yes_target = inserted_ids[idx + 1] if idx + 1 < len(inserted_ids) else None
                    conn.execute(
                        "UPDATE workflow_steps SET yes_step_id = ? WHERE id = ?",
                        (yes_target, step_id),
                    )
            break
        except sqlite3.OperationalError as exc:
            if attempt == 2 or "workflow_processes" not in str(exc):
                raise
            init_db()


def fetch_workflows() -> pd.DataFrame:
    init_db()
    query = """
        SELECT
            w.id AS workflow_id,
            w.name AS workflow_name,
            w.owner AS workflow_owner,
            w.description AS workflow_description,
            w.created_at AS workflow_created_at,
            s.id AS step_id,
            s.step_order,
            s.title AS step_title,
            s.status AS step_status,
            s.yes_step_id,
            s.no_step_id,
            s.updated_at AS step_updated_at
        FROM workflow_processes w
        LEFT JOIN workflow_steps s ON s.workflow_id = w.id
        ORDER BY w.created_at DESC, s.step_order ASC
    """

    try:
        with get_connection() as conn:
            df = pd.read_sql_query(query, conn)
    except (sqlite3.OperationalError, pd.errors.DatabaseError) as exc:
        if "workflow_processes" in str(exc):
            return pd.DataFrame(
                columns=[
                    "workflow_id",
                    "workflow_name",
                    "workflow_owner",
                    "workflow_description",
                    "workflow_created_at",
                    "step_id",
                    "step_order",
                    "step_title",
                    "step_status",
                    "yes_step_id",
                    "no_step_id",
                    "step_updated_at",
                ]
            )
        raise

    if not df.empty:
        df["workflow_created_at"] = pd.to_datetime(df["workflow_created_at"])
        if "step_updated_at" in df.columns:
            df["step_updated_at"] = pd.to_datetime(df["step_updated_at"])

    return df


@st.cache_data(show_spinner=False)
def load_workflows() -> pd.DataFrame:
    return fetch_workflows()


def clear_cached_workflows() -> None:
    load_workflows.clear()  # type: ignore[attr-defined]


def trigger_rerun() -> None:
    rerun = getattr(st, "rerun", None)
    if callable(rerun):
        rerun()
        return

    experimental_rerun = getattr(st, "experimental_rerun", None)
    if callable(experimental_rerun):
        experimental_rerun()


def update_workflow_steps(step_updates: Dict[int, Dict[str, object]]) -> None:
    if not step_updates:
        return

    init_db()
    now_iso = dt.datetime.utcnow().isoformat(timespec="seconds")
    for attempt in (1, 2):
        try:
            with get_connection() as conn:
                ensure_workflow_tables(conn)
                for step_id, fields in step_updates.items():
                    if not fields:
                        continue

                    columns: List[str] = []
                    values: List[object] = []

                    for column in ("title", "step_order", "yes_step_id", "no_step_id"):
                        if column in fields:
                            if column == "title" and isinstance(fields[column], str):
                                values.append(fields[column].strip())
                            else:
                                values.append(fields[column])
                            columns.append(f"{column} = ?")

                    if not columns:
                        continue

                    columns.append("updated_at = ?")
                    values.append(now_iso)
                    values.append(step_id)

                    conn.execute(
                        f"UPDATE workflow_steps SET {', '.join(columns)} WHERE id = ?",
                        values,
                    )
            break
        except sqlite3.OperationalError as exc:
            if attempt == 2 or "workflow_steps" not in str(exc):
                raise
            init_db()


def insert_workflow_step(
    workflow_id: int,
    step_order: int,
    title: str,
    yes_step_id: Optional[int],
    no_step_id: Optional[int],
) -> int:
    init_db()
    now_iso = dt.datetime.utcnow().isoformat(timespec="seconds")
    payload = (
        int(workflow_id),
        int(step_order),
        title.strip(),
        "Not Started",
        yes_step_id,
        no_step_id,
        now_iso,
        now_iso,
    )

    for attempt in (1, 2):
        try:
            with get_connection() as conn:
                ensure_workflow_tables(conn)
                cursor = conn.execute(
                    """
                    INSERT INTO workflow_steps (
                        workflow_id, step_order, title, status, yes_step_id, no_step_id, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    payload,
                )
                return int(cursor.lastrowid)
        except sqlite3.OperationalError as exc:
            if attempt == 2 or "workflow_steps" not in str(exc):
                raise
            init_db()
    return 0


def delete_workflow_steps(step_ids: Iterable[int]) -> None:
    ids = [int(step_id) for step_id in step_ids]
    if not ids:
        return

    init_db()
    for attempt in (1, 2):
        try:
            with get_connection() as conn:
                ensure_workflow_tables(conn)
                placeholders = ",".join("?" for _ in ids)
                conn.execute(
                    f"DELETE FROM workflow_steps WHERE id IN ({placeholders})",
                    ids,
                )
            break
        except sqlite3.OperationalError as exc:
            if attempt == 2 or "workflow_steps" not in str(exc):
                raise
            init_db()


def delete_workflow_process(workflow_id: int) -> None:
    init_db()
    for attempt in (1, 2):
        try:
            with get_connection() as conn:
                ensure_workflow_tables(conn)
                conn.execute(
                    "DELETE FROM workflow_processes WHERE id = ?",
                    (int(workflow_id),),
                )
            break
        except sqlite3.OperationalError as exc:
            if attempt == 2 or "workflow_processes" not in str(exc):
                raise
            init_db()


def render_workflow_diagram(steps_df: pd.DataFrame) -> None:
    graphviz_code = build_graphviz_workflow(steps_df)
    if graphviz_code:
        st.graphviz_chart(graphviz_code, use_container_width=True)


def render_workflows() -> None:
    st.subheader("🧭 Process workflow builder")
    st.markdown(
        "Establish repeatable handover steps and track their progress visually.")

    with st.form("create_workflow_form", clear_on_submit=True):
        name = st.text_input("Workflow name", placeholder="e.g., New manager onboarding")
        owner = st.text_input("Owner (optional)", placeholder="Who leads this process?")
        description = st.text_area(
            "Description",
            placeholder="What is the goal and scope of this workflow?",
        )
        steps_text = st.text_area(
            "Workflow steps (one per line)",
            placeholder="Kick-off alignment\nTool access handover\nShadowing and sign-off",
        )
        create_submitted = st.form_submit_button("Create workflow", type="primary")

    if create_submitted:
        steps = [line.strip() for line in steps_text.splitlines() if line.strip()]
        if not name.strip():
            st.error("Workflow name is required.")
        elif not steps:
            st.error("Add at least one workflow step.")
        else:
            add_workflow_process(name, owner, description, steps)
            clear_cached_workflows()
            st.success("Workflow created.")
            trigger_rerun()

    df = load_workflows()

    if df.empty or df["workflow_id"].isna().all():
        st.info("No workflows captured yet. Use the form above to build your first one.")
        return

    with st.expander("Manage workflows", expanded=False):
        unique_workflows = (
            df.dropna(subset=["workflow_id"]).sort_values("workflow_created_at", ascending=False)
        )
        workflow_choices = {}
        for row in unique_workflows.itertuples():
            label_bits = [row.workflow_name]
            if isinstance(row.workflow_owner, str) and row.workflow_owner.strip():
                label_bits.append(f"Owner: {row.workflow_owner.strip()}")
            created_at = getattr(row, "workflow_created_at", None)
            if pd.notna(created_at):
                label_bits.append(created_at.strftime("%d %b %Y"))
            label = " · ".join(label_bits)
            workflow_choices[label] = int(row.workflow_id)

        if workflow_choices:
            col_pick, col_delete = st.columns([3, 1])
            with col_pick:
                selected_label = st.selectbox(
                    "Select a workflow to delete",
                    options=list(workflow_choices.keys()),
                    key="workflow_delete_select",
                )
            with col_delete:
                if st.button("Delete", key="workflow_delete_button"):
                    delete_workflow_process(workflow_choices[selected_label])
                    clear_cached_workflows()
                    st.success("Workflow removed.")
                    trigger_rerun()
        else:
            st.info("No workflows available to delete.")

    grouped = df.groupby("workflow_id", sort=False)
    for workflow_id, group in grouped:
        workflow = group.iloc[0]
        steps_df = group.dropna(subset=["step_id"]).sort_values("step_order")

        header = workflow["workflow_name"]
        meta_bits: List[str] = []
        if isinstance(workflow.get("workflow_owner"), str) and workflow["workflow_owner"].strip():
            meta_bits.append(f"Owner: {workflow['workflow_owner'].strip()}")
        created_at = workflow.get("workflow_created_at")
        if pd.notna(created_at):
            meta_bits.append(f"Created {created_at.strftime('%d %b %Y')}")

        expander_label = header
        if meta_bits:
            expander_label = f"{header} · {' · '.join(meta_bits)}"

        with st.expander(expander_label, expanded=False):
            description = workflow.get("workflow_description")
            if isinstance(description, str) and description.strip():
                st.markdown(description.strip())

            tab_visual, tab_barfi = st.tabs([
                "Diagram & table",
                "Barfi editor (beta)",
            ])

            with tab_visual:
                if steps_df.empty:
                    st.info("No steps logged for this workflow yet. Use the table below to add steps.")
                else:
                    render_workflow_diagram(steps_df)

                base_columns = ["step_id", "step_order", "step_title", "yes_step_id", "no_step_id"]
                if steps_df.empty:
                    table_source = pd.DataFrame(columns=base_columns)
                else:
                    table_source = steps_df[base_columns].copy()

                table_source = table_source.rename(
                    columns={
                        "step_id": "step_id",
                        "step_order": "step_order",
                        "step_title": "title",
                        "yes_step_id": "yes_step_id",
                        "no_step_id": "no_step_id",
                    }
                )

                if "step_id" in table_source.columns:
                    table_source["step_id"] = table_source["step_id"].astype("Int64")

                row_count = max(len(table_source), 1)
                table_height = min(700, 60 * row_count + 120)

                title_len = int(table_source["title"].astype(str).str.len().max()) if not table_source.empty else 0
                step_len = int(table_source["step_order"].astype(str).str.len().max()) if not table_source.empty else 1
                yes_len = int(table_source["yes_step_id"].astype(str).str.len().max()) if not table_source.empty else 1
                no_len = int(table_source["no_step_id"].astype(str).str.len().max()) if not table_source.empty else 1

                def _pick_width(length: int) -> str:
                    if length >= 25:
                        return "large"
                    if length >= 12:
                        return "medium"
                    return "small"

                editor_df = st.data_editor(
                    table_source,
                    num_rows="dynamic",
                    use_container_width=True,
                    height=table_height,
                    key=f"workflow_table_editor_{workflow_id}",
                    column_config={
                        "step_id": st.column_config.NumberColumn("Step ID", disabled=True, width="small"),
                        "step_order": st.column_config.NumberColumn(
                            "Step order", min_value=1, width=_pick_width(step_len)
                        ),
                        "title": st.column_config.TextColumn("Title", width=_pick_width(title_len)),
                        "yes_step_id": st.column_config.NumberColumn(
                            "Yes → Step ID", min_value=1, required=False, width=_pick_width(yes_len)
                        ),
                        "no_step_id": st.column_config.NumberColumn(
                            "No → Step ID", min_value=1, required=False, width=_pick_width(no_len)
                        ),
                    },
                    hide_index=True,
                )

                if st.button("Save table changes", key=f"workflow_table_save_{workflow_id}"):
                    if editor_df.empty:
                        st.error("Add at least one step before saving.")
                    else:
                        try:
                            valid_step_ids: Set[int] = set(
                                table_source["step_id"].dropna().astype(int).tolist()
                            )
                            original_ids = set(
                                table_source["step_id"].dropna().astype(int).tolist()
                            )
                            edited_ids = set(
                                editor_df["step_id"].dropna().astype(int).tolist()
                            )
                            to_delete = original_ids - edited_ids

                            updates: Dict[int, Dict[str, object]] = {}
                            inserts: List[Dict[str, object]] = []

                            for row in editor_df.itertuples(index=False):
                                step_id = row.step_id
                                title_raw = row.title
                                order_raw = row.step_order

                                if not isinstance(title_raw, str) or not title_raw.strip():
                                    raise ValueError("Every step must have a title.")
                                if pd.isna(order_raw):
                                    raise ValueError("Every step must have an order number.")

                                yes_raw = row.yes_step_id
                                no_raw = row.no_step_id

                                yes_val = None if pd.isna(yes_raw) else int(yes_raw)
                                no_val = None if pd.isna(no_raw) else int(no_raw)
                                for value, label in ((yes_val, "Yes"), (no_val, "No")):
                                    if value is not None and value not in valid_step_ids:
                                        raise ValueError(
                                            f"{label} → Step ID {value} does not exist yet. "
                                            "Use one of the Step ID values listed in the table."
                                        )
                                payload = {
                                    "step_order": int(order_raw),
                                    "title": title_raw.strip(),
                                    "yes_step_id": yes_val,
                                    "no_step_id": no_val,
                                }

                                if pd.isna(step_id):
                                    inserts.append(payload)
                                else:
                                    updates[int(step_id)] = payload

                            if to_delete:
                                delete_workflow_steps(to_delete)

                            if updates:
                                update_workflow_steps(updates)

                            workflow_id_int = int(workflow_id)
                            for payload in inserts:
                                insert_workflow_step(
                                    workflow_id_int,
                                    payload["step_order"],
                                    payload["title"],
                                    payload["yes_step_id"],
                                    payload["no_step_id"],
                                )

                            clear_cached_workflows()
                            st.success("Workflow table updated.")
                            trigger_rerun()
                        except ValueError as exc:
                            st.error(str(exc))

            with tab_barfi:
                render_barfi_editor(expander_label, int(workflow_id))


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
    show_excel = EXCEL_ENGINE is not None
    columns = 3 if show_excel else 2
    layout = st.columns(columns)

    col_csv = layout[0]
    col_summary = layout[-1]
    excel_bytes = None

    if show_excel:
        col_excel = layout[1]
        try:
            excel_bytes = build_excel_report(df)
        except RuntimeError:
            show_excel = False

    summary_markdown = build_summary(df)

    with col_csv:
        st.download_button(
            label="⬇️ CSV",
            data=csv,
            file_name="handover_topics.csv",
            mime="text/csv",
            use_container_width=True,
        )

    if show_excel and excel_bytes:
        with col_excel:  # type: ignore[name-defined]
            st.download_button(
                label="📊 Excel report",
                data=excel_bytes,
                file_name="handover_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    elif not show_excel:
        st.info(
            "Install `xlsxwriter` or `openpyxl` to enable Excel exports.",
            icon="ℹ️",
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
    st.markdown(
        """
        <style>
            :root {
                font-size: 18px;
            }
            body, .block-container {
                font-size: 1.1rem !important;
            }
            h1, h2, h3, h4, h5, h6 {
                font-weight: 600;
                line-height: 1.35;
            }
            .stMarkdown p,
            .stMarkdown li,
            div[data-testid="stMetricValue"],
            div[data-testid="stMetricLabel"],
            .stDataFrame,
            label,
            .stTextInput input,
            .stTextArea textarea,
            div[data-baseweb="select"],
            button,
            .stButton button {
                font-size: 1.15rem !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("Handover planning dashboard")
    st.caption("Track conversations, owners, and next steps ahead of your transition.")

    init_db()

    tab_add, tab_review, tab_workflow = st.tabs(
        ["Add topic", "Review & update", "Workflow builder"]
    )

    with tab_add:
        render_add_topic()

    with tab_review:
        render_review()

    with tab_workflow:
        render_workflows()


if __name__ == "__main__":
    main()

