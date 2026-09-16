from __future__ import annotations

import csv
import io
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document
from openai import OpenAI
from openpyxl import load_workbook
from pypdf import PdfReader


APP_TITLE = "Project Update Assistant"
SAMPLES_DIR = Path(__file__).parent / "samples"
SUPPORTED_REQUIREMENTS = ["xlsx", "xls", "docx", "pdf", "txt"]
SUPPORTED_NOTES = ["pdf", "docx", "txt"]

CATEGORY_LABELS = {
    "suggested_requirements_updates": "Suggested requirements updates",
    "information_requiring_validation": "Information requiring validation",
    "conflicts_or_contradictions": "Conflicts or contradictions",
    "requirements_already_captured": "Requirements already captured",
    "open_questions": "Open questions",
    "draft_actions": "Draft actions",
}

SAMPLE_REQUIREMENTS = """Project: Northstar client portal
Owner: Delivery team
Status: In discovery

Release plan
- Target launch: 14 June 2026
- The first release includes account setup, project overview, and weekly status exports.

Reporting
- Weekly status exports are delivered as CSV to the client operations contact.
- The portal should not send messages automatically.
"""

SAMPLE_MEETING_NOTES = """Northstar discovery workshop — 28 May 2026

The client said the launch window may need to move to 5 July 2026 because legal review
is taking longer than expected. Priya said she will confirm the date after the legal
review on 3 June. This is not a final approval yet.

The client asked whether weekly status exports could also be available as PDF. The
team agreed to investigate, but no format decision was made.

The client confirmed that the portal must not send messages automatically. The client
also asked who will own the operations handover.
"""


@dataclass
class DocumentContent:
    filename: str
    file_type: str
    text: str
    sections: list[dict[str, str]]
    source: str = "upload"


def inject_styles() -> None:
    """Small, user-requested theme layer for the internal-tool interface."""
    st.markdown(
        """
        <style>
        :root {
            --navy: #102a43;
            --navy-deep: #071b2f;
            --teal: #0f766e;
            --teal-soft: #d9f3ef;
            --ink: #243b53;
            --muted: #627d98;
            --line: #d9e2ec;
            --surface: #ffffff;
            --canvas: #f4f7fa;
        }
        .stApp { background: var(--canvas); color: var(--ink); }
        .block-container { max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem; }
        [data-testid="stSidebar"] { background: var(--navy-deep); }
        [data-testid="stSidebar"] * { color: #e6f0f5 !important; }
        .app-kicker { color: #5eead4; font-weight: 700; letter-spacing: .12em; font-size: .73rem; text-transform: uppercase; }
        .app-title { color: var(--navy); font-size: 2.7rem; font-weight: 800; line-height: 1.05; margin: .25rem 0 .6rem; }
        .app-subtitle { color: var(--muted); font-size: 1.05rem; max-width: 720px; margin-bottom: 1.4rem; }
        .step-card { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 1rem 1.1rem; min-height: 94px; }
        .step-number { color: var(--teal); font-size: .72rem; font-weight: 800; letter-spacing: .1em; text-transform: uppercase; }
        .step-title { color: var(--navy); font-weight: 700; margin-top: .35rem; }
        .step-copy { color: var(--muted); font-size: .84rem; margin-top: .2rem; }
        .section-label { color: var(--navy); font-size: 1.25rem; font-weight: 750; margin-top: 1.35rem; }
        .finding-card { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 1rem 1.1rem; margin: .6rem 0 1rem; }
        .finding-title { color: var(--navy); font-weight: 750; font-size: 1.02rem; }
        .finding-label { color: var(--muted); text-transform: uppercase; letter-spacing: .08em; font-size: .68rem; font-weight: 800; margin-top: .85rem; }
        .finding-value { color: var(--ink); font-size: .92rem; line-height: 1.45; }
        .confidence { display: inline-block; border-radius: 999px; padding: .22rem .55rem; background: var(--teal-soft); color: #115e59; font-size: .72rem; font-weight: 800; }
        .notice { border-left: 4px solid var(--teal); background: #e8f7f5; padding: .85rem 1rem; border-radius: 0 8px 8px 0; color: var(--ink); }
        .warning { border-left-color: #d97706; background: #fff7ed; }
        .doc-chip { display: inline-block; border: 1px solid var(--line); background: var(--surface); border-radius: 999px; padding: .3rem .7rem; color: var(--ink); font-size: .8rem; margin: .2rem .35rem .2rem 0; }
        div[data-testid="stFileUploader"] { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: .3rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def file_type(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip(".")


def non_empty_rows(rows: list[str]) -> list[str]:
    return [row.strip() for row in rows if row and row.strip()]


def extract_text_file(raw: bytes) -> tuple[str, list[dict[str, str]]]:
    text = raw.decode("utf-8", errors="replace").strip()
    sections = [{"title": "Text content", "content": text}] if text else []
    return text, sections


def extract_docx(raw: bytes) -> tuple[str, list[dict[str, str]]]:
    document = Document(io.BytesIO(raw))
    sections: list[dict[str, str]] = []
    paragraphs = non_empty_rows([paragraph.text for paragraph in document.paragraphs])
    if paragraphs:
        sections.append({"title": "Paragraphs", "content": "\n".join(paragraphs)})

    for index, table in enumerate(document.tables, start=1):
        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                rows.append(" | ".join(cells))
        if rows:
            sections.append({"title": f"Table {index}", "content": "\n".join(rows)})

    text = "\n\n".join(f"{section['title']}:\n{section['content']}" for section in sections)
    return text, sections


def extract_pdf(raw: bytes) -> tuple[str, list[dict[str, str]]]:
    reader = PdfReader(io.BytesIO(raw))
    sections: list[dict[str, str]] = []
    for index, page in enumerate(reader.pages, start=1):
        content = (page.extract_text() or "").strip()
        if content:
            sections.append({"title": f"Page {index}", "content": content})
    text = "\n\n".join(f"{section['title']}:\n{section['content']}" for section in sections)
    return text, sections


def is_populated(value: Any) -> bool:
    return value is not None and str(value).strip() != ""


def compact_spreadsheet_rows(rows: list[list[Any]]) -> list[list[str]]:
    """Remove wholly empty rows and columns while preserving cell order."""
    populated_rows = [row for row in rows if any(is_populated(value) for value in row)]
    if not populated_rows:
        return []
    column_count = max(len(row) for row in populated_rows)
    populated_columns = [
        column_index
        for column_index in range(column_count)
        if any(
            column_index < len(row) and is_populated(row[column_index])
            for row in populated_rows
        )
    ]
    return [
        [
            "" if column_index >= len(row) or row[column_index] is None else str(row[column_index]).strip()
            for column_index in populated_columns
        ]
        for row in populated_rows
    ]


def spreadsheet_sections(sheet_rows: list[tuple[str, list[list[Any]]]]) -> list[dict[str, str]]:
    sections: list[dict[str, str]] = []
    for sheet_name, rows in sheet_rows:
        compact_rows = compact_spreadsheet_rows(rows)
        if compact_rows:
            content = "\n".join(" | ".join(row) for row in compact_rows)
            sections.append({"title": f"Sheet: {sheet_name}", "content": content})
    return sections


def extract_xlsx(raw: bytes) -> tuple[str, list[dict[str, str]]]:
    workbook = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    visible_sheets = [
        sheet
        for sheet in workbook.worksheets
        if sheet.sheet_state == "visible"
    ]
    if not visible_sheets:
        raise ValueError("The workbook has no visible worksheets.")
    sheet_rows = [
        (sheet.title, [list(values) for values in sheet.iter_rows(values_only=True)])
        for sheet in visible_sheets
    ]
    sections = spreadsheet_sections(sheet_rows)
    if not sections:
        raise ValueError("The workbook has no populated cells in its visible worksheets.")
    text = "\n\n".join(f"{section['title']}:\n{section['content']}" for section in sections)
    return text, sections


def extract_xls(raw: bytes) -> tuple[str, list[dict[str, str]]]:
    workbook = pd.ExcelFile(io.BytesIO(raw), engine="xlrd")
    sheet_rows = []
    for sheet_name in workbook.sheet_names:
        dataframe = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=None,
            dtype=object,
        )
        sheet_rows.append((sheet_name, dataframe.where(dataframe.notna(), None).values.tolist()))
    sections = spreadsheet_sections(sheet_rows)
    if not sections:
        raise ValueError("The workbook has no populated cells in its worksheets.")
    text = "\n\n".join(f"{section['title']}:\n{section['content']}" for section in sections)
    return text, sections


def extract_document(filename: str, raw: bytes, source: str = "upload") -> DocumentContent:
    kind = file_type(filename)
    if kind == "txt":
        text, sections = extract_text_file(raw)
    elif kind == "docx":
        text, sections = extract_docx(raw)
    elif kind == "pdf":
        text, sections = extract_pdf(raw)
    elif kind == "xlsx":
        text, sections = extract_xlsx(raw)
    elif kind == "xls":
        text, sections = extract_xls(raw)
    else:
        raise ValueError(f"Unsupported file type: .{kind or 'unknown'}")

    if not text.strip():
        raise ValueError(f"No readable text was found in {filename}.")
    return DocumentContent(filename, kind.upper(), text, sections, source)


def sample_document(kind: str) -> DocumentContent:
    if kind == "requirements":
        return extract_document("northstar-requirements.txt", SAMPLE_REQUIREMENTS.encode(), "sample")
    return extract_document("northstar-meeting-notes.txt", SAMPLE_MEETING_NOTES.encode(), "sample")


def truncate(text: str, limit: int = 2800) -> str:
    return text if len(text) <= limit else f"{text[:limit].rstrip()}…"


def evidence_sentence(text: str, keywords: list[str]) -> str:
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    for sentence in sentences:
        if any(keyword.lower() in sentence.lower() for keyword in keywords if keyword):
            return sentence.strip()
    return sentences[0].strip() if sentences else text[:300]


def finding(
    title: str,
    evidence: str,
    current: str,
    changed: str,
    handling: str,
    confidence: str,
) -> dict[str, str]:
    return {
        "title": title,
        "meeting_evidence": evidence or "No direct meeting evidence found.",
        "current_document_content": current or "No corresponding content found.",
        "what_changed": changed,
        "suggested_handling": handling,
        "confidence": confidence if confidence in {"High", "Medium", "Low"} else "Low",
    }


def local_compare(requirements: DocumentContent, notes: DocumentContent) -> dict[str, Any]:
    """A transparent, evidence-based fallback for demos when OpenAI is not configured."""
    requirement_text = requirements.text
    meeting_text = notes.text
    result: dict[str, Any] = {key: [] for key in CATEGORY_LABELS}

    req_dates = set(re.findall(r"\b(?:\d{1,2}\s+)?(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,\s*\d{4})?|\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", requirement_text, flags=re.I))
    note_dates = set(re.findall(r"\b(?:\d{1,2}\s+)?(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,\s*\d{4})?|\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", meeting_text, flags=re.I))
    new_dates = sorted(note_dates - req_dates)
    if new_dates:
        evidence = evidence_sentence(meeting_text, new_dates)
        current = evidence_sentence(requirement_text, ["launch", "target", "release", "date"])
        result["suggested_requirements_updates"].append(
            finding(
                "Review the launch date before updating the requirements",
                evidence,
                current,
                f"The meeting mentions date(s) not present in the current document: {', '.join(new_dates)}.",
                "Validate the date with the client, then update the requirements only if it is explicitly confirmed.",
                "Medium",
            )
        )

    if re.search(r"\b(?:approved|agreed|decided|confirmed)\b", meeting_text, re.I):
        evidence = evidence_sentence(meeting_text, ["approved", "agreed", "decided", "confirmed"])
        result["information_requiring_validation"].append(
            finding(
                "Treat meeting language as a proposal until the decision is confirmed",
                evidence,
                "The requirements document does not record a confirmed change for this statement.",
                "The notes contain decision-like language alongside uncertainty or follow-up ownership.",
                "Ask the client to confirm the decision in writing before treating it as a requirement.",
                "Medium",
            )
        )

    if "must not send messages automatically" in requirement_text.lower() and "must not send messages automatically" in meeting_text.lower():
        evidence = evidence_sentence(meeting_text, ["must not send messages automatically"])
        current = evidence_sentence(requirement_text, ["must not send messages automatically"])
        result["requirements_already_captured"].append(
            finding(
                "No change found: automatic messaging restriction is already captured",
                evidence,
                current,
                "The meeting repeats an existing restriction.",
                "Keep the current requirement and record the repeated confirmation in project notes.",
                "High",
            )
        )

    question_lines = [
        line.strip()
        for line in meeting_text.splitlines()
        if "?" in line or re.search(r"\b(?:who|which|what|when|how)\b", line, re.I)
    ]
    for line in question_lines[:3]:
        result["open_questions"].append(
            finding(
                line.rstrip("?") or "Open question from the meeting",
                line,
                "No answer is present in the current requirements document.",
                "The meeting raises a question that is not resolved in the current document.",
                "Add an owner for follow-up only after the team or client explicitly assigns one.",
                "Low",
            )
        )

    owner_match = re.search(r"\b([A-Z][a-z]+)\s+(?:said|will|owns|is responsible)\b", meeting_text)
    deadline_match = re.search(r"\b(?:by|after|on)\s+(\d{1,2}\s+[A-Za-z]+\b|\d{1,2}\s+[A-Za-z]+\s+\d{4})", meeting_text)
    if owner_match and deadline_match:
        result["draft_actions"].append(
            finding(
                "Draft follow-up action mentioned in the notes",
                evidence_sentence(meeting_text, [owner_match.group(1), deadline_match.group(1)]),
                "No matching action is listed in the current requirements document.",
                "The notes explicitly associate a person and timing with a follow-up.",
                "Keep this as a draft action until the owner and deadline are confirmed.",
                "Medium",
            )
            | {"owner": owner_match.group(1), "deadline": deadline_match.group(1)}
        )

    result["draft_internal_slack"] = (
        "Draft internal update\n\n"
        "The latest meeting notes have been compared with the current requirements. "
        "Please review the flagged evidence before sharing any change externally. "
        "No requirement has been updated automatically."
    )
    result["draft_client_follow_up"] = (
        "Draft client follow-up\n\n"
        "Thank you for the latest discussion. Could you please confirm the points "
        "flagged for validation, including any proposed launch-date change and the "
        "open reporting question?"
    )
    result["draft_asana_tasks"] = (
        "Draft Asana tasks\n\n"
        "1. Validate proposed requirement changes against the meeting evidence.\n"
        "2. Resolve open questions before updating the requirements document."
    )
    return result


def comparison_prompt(requirements: DocumentContent, notes: DocumentContent) -> str:
    schema = {
        "suggested_requirements_updates": ["finding objects"],
        "information_requiring_validation": ["finding objects"],
        "conflicts_or_contradictions": ["finding objects"],
        "requirements_already_captured": ["finding objects"],
        "open_questions": ["finding objects"],
        "draft_actions": ["finding objects with owner and deadline only if explicit"],
        "draft_internal_slack": "plain text draft",
        "draft_client_follow_up": "plain text draft",
        "draft_asana_tasks": "plain text draft",
    }
    return f"""
You are reviewing two project documents for a consultant. Compare them carefully and return JSON only.
Never treat a suggestion as a confirmed client decision. Never invent owners, deadlines,
requirements, technical solutions, or missing content. Clearly label uncertainty.
Every finding must include: title, meeting_evidence, current_document_content, what_changed,
suggested_handling, confidence (High, Medium, or Low). Draft actions may include owner and
deadline only when they are explicitly stated in the meeting notes. Keep all drafts clearly
marked as drafts. If a category has no findings, return an empty array.

Required JSON shape:
{json.dumps(schema, indent=2)}

CURRENT REQUIREMENTS DOCUMENT ({requirements.filename}):
{truncate(requirements.text, 18000)}

LATEST MEETING NOTES ({notes.filename}):
{truncate(notes.text, 18000)}
"""


def normalize_finding(item: Any) -> dict[str, str]:
    if not isinstance(item, dict):
        return finding(str(item), "", "", "", "Review the source documents.", "Low")
    return finding(
        str(item.get("title", "Untitled finding")),
        str(item.get("meeting_evidence", "")),
        str(item.get("current_document_content", "")),
        str(item.get("what_changed", "")),
        str(item.get("suggested_handling", "")),
        str(item.get("confidence", "Low")),
    ) | {
        "owner": str(item.get("owner", "")),
        "deadline": str(item.get("deadline", "")),
    }


def normalize_result(raw: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key in CATEGORY_LABELS:
        values = raw.get(key, [])
        result[key] = [normalize_finding(value) for value in values] if isinstance(values, list) else []
    for key in ("draft_internal_slack", "draft_client_follow_up", "draft_asana_tasks"):
        result[key] = str(raw.get(key, "No draft was generated."))
    return result


def run_ai_comparison(requirements: DocumentContent, notes: DocumentContent) -> tuple[dict[str, Any] | None, str | None]:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None, None
    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            temperature=0.1,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You produce careful, evidence-based project document comparisons."},
                {"role": "user", "content": comparison_prompt(requirements, notes)},
            ],
        )
        content = response.choices[0].message.content or "{}"
        return normalize_result(json.loads(content)), None
    except Exception as error:
        return None, f"Live comparison was unavailable ({type(error).__name__}). The transparent local comparison is shown instead."


def get_document_from_upload(uploaded: Any, fallback: DocumentContent | None) -> DocumentContent | None:
    if uploaded is not None:
        return extract_document(uploaded.name, uploaded.getvalue())
    return fallback


def render_document_preview(document: DocumentContent, label: str) -> None:
    st.markdown(f"#### {label}")
    st.markdown(f'<span class="doc-chip">{document.filename}</span><span class="doc-chip">{document.file_type}</span>', unsafe_allow_html=True)
    is_excel = document.file_type in {"XLSX", "XLS"}
    with st.expander("Excel preview" if is_excel else "Show extracted sections", expanded=is_excel):
        for section in document.sections:
            st.markdown(f"**{section['title']}**")
            st.code(truncate(section["content"], 2400), language="text")


def render_finding(category: str, item: dict[str, str], index: int) -> None:
    title = item.get("title", "Untitled finding")
    confidence = item.get("confidence", "Low")
    st.markdown('<div class="finding-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="finding-title">{title}</div><span class="confidence">{confidence} confidence</span>', unsafe_allow_html=True)
    columns = st.columns(2)
    fields = [
        ("Meeting evidence", item.get("meeting_evidence", "")),
        ("Current document content", item.get("current_document_content", "")),
        ("What changed", item.get("what_changed", "")),
        ("Suggested handling", item.get("suggested_handling", "")),
    ]
    for column, (field_label, value) in zip(columns * 2, fields):
        with column:
            st.markdown(f'<div class="finding-label">{field_label}</div><div class="finding-value">{value}</div>', unsafe_allow_html=True)
    if category == "draft_actions":
        owner = item.get("owner") or "Not explicitly stated"
        deadline = item.get("deadline") or "Not explicitly stated"
        st.caption(f"Owner: {owner} · Deadline: {deadline}")
    st.checkbox("Approve for export", key=f"approve_{category}_{index}")
    st.markdown("</div>", unsafe_allow_html=True)


def approved_findings(result: dict[str, Any]) -> list[dict[str, str]]:
    approved = []
    for category in CATEGORY_LABELS:
        for index, item in enumerate(result.get(category, [])):
            if st.session_state.get(f"approve_{category}_{index}", False):
                approved.append({"category": CATEGORY_LABELS[category], **item})
    return approved


def export_csv(findings: list[dict[str, str]]) -> str:
    output = io.StringIO()
    columns = ["category", "title", "confidence", "meeting_evidence", "current_document_content", "what_changed", "suggested_handling", "owner", "deadline"]
    writer = csv.DictWriter(output, fieldnames=columns)
    writer.writeheader()
    for item in findings:
        writer.writerow({column: item.get(column, "") for column in columns})
    return output.getvalue()


def export_txt(findings: list[dict[str, str]]) -> str:
    blocks = []
    for item in findings:
        blocks.append(
            "\n".join(
                [
                    f"[{item.get('category', 'Finding')}] {item.get('title', '')}",
                    f"Confidence: {item.get('confidence', '')}",
                    f"Meeting evidence: {item.get('meeting_evidence', '')}",
                    f"Current document content: {item.get('current_document_content', '')}",
                    f"What changed: {item.get('what_changed', '')}",
                    f"Suggested handling: {item.get('suggested_handling', '')}",
                ]
            )
        )
    return "\n\n".join(blocks)


def render_results(result: dict[str, Any]) -> None:
    st.markdown('<div class="section-label">Review findings</div>', unsafe_allow_html=True)
    st.caption("Suggestions are not confirmed decisions. Approve individual findings only after human review.")
    for category, label in CATEGORY_LABELS.items():
        items = result.get(category, [])
        with st.expander(f"{label} · {len(items)}", expanded=bool(items)):
            if not items:
                st.caption("No findings in this category.")
            for index, item in enumerate(items):
                render_finding(category, item, index)

    st.markdown('<div class="section-label">Draft communications</div>', unsafe_allow_html=True)
    st.caption("These are drafts for review only. The app does not send messages or create tasks.")
    draft_columns = st.columns(3)
    for column, label, key in zip(
        draft_columns,
        ["Internal Slack update", "Client follow-up", "Asana tasks"],
        ["draft_internal_slack", "draft_client_follow_up", "draft_asana_tasks"],
    ):
        with column:
            st.markdown(f"**{label}**")
            st.text_area(label, result.get(key, ""), height=190, disabled=True, label_visibility="collapsed")

    approved = approved_findings(result)
    st.markdown('<div class="section-label">Export approved findings</div>', unsafe_allow_html=True)
    if approved:
        st.success(f"{len(approved)} approved finding(s) ready for export.")
        export_columns = st.columns(2)
        with export_columns[0]:
            st.download_button("Download approved CSV", export_csv(approved), "approved-findings.csv", "text/csv", use_container_width=True)
        with export_columns[1]:
            st.download_button("Download approved TXT", export_txt(approved), "approved-findings.txt", "text/plain", use_container_width=True)
    else:
        st.info("Approve at least one finding to enable exports.")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="◈", layout="wide", initial_sidebar_state="expanded")
    inject_styles()

    with st.sidebar:
        st.markdown("## Project Update Assistant")
        st.caption("Evidence-first document review for consulting teams.")
        st.divider()
        if os.getenv("OPENAI_API_KEY"):
            st.success("Live AI comparison enabled")
        else:
            st.warning("Demo comparison mode\n\nAdd OPENAI_API_KEY in Replit Secrets for live OpenAI comparison.")
        st.divider()
        st.markdown("### Upload policy")
        st.caption("Only upload anonymised or approved documents. Files are processed in memory for this session and are not permanently stored.")
        st.caption(f"Session started {datetime.now().strftime('%d %b %Y')}")

    st.markdown('<div class="app-kicker">Internal review workspace</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-title">Project Update Assistant</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-subtitle">Compare the latest project conversation with the current requirements, then review evidence-backed changes before anything reaches a client or delivery system.</div>', unsafe_allow_html=True)

    steps = st.columns(4)
    for column, number, title, copy in [
        (steps[0], "01", "Upload", "Add requirements and meeting notes"),
        (steps[1], "02", "Confirm", "Check extracted sections"),
        (steps[2], "03", "Compare", "Find changes and gaps"),
        (steps[3], "04", "Approve", "Export reviewed findings"),
    ]:
        with column:
            st.markdown(f'<div class="step-card"><div class="step-number">{number}</div><div class="step-title">{title}</div><div class="step-copy">{copy}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">1. Add your project documents</div>', unsafe_allow_html=True)
    st.markdown('<div class="notice warning"><strong>Confidentiality reminder:</strong> only upload anonymised or approved documents. Uploaded content is processed in memory for this session and is not automatically written back to your files.</div>', unsafe_allow_html=True)
    demo_col, status_col = st.columns([1, 2])
    with demo_col:
        if st.button("Load anonymised example", use_container_width=True):
            st.session_state["demo_loaded"] = True
            st.session_state.pop("comparison", None)
            st.rerun()
    with status_col:
        st.caption("Use the example to see the complete workflow without confidential client data.")

    uploader_columns = st.columns(2)
    with uploader_columns[0]:
        requirements_upload = st.file_uploader("Current requirements document", type=SUPPORTED_REQUIREMENTS, help="Supported: XLSX, XLS, DOCX, PDF, TXT", key="requirements_upload")
    with uploader_columns[1]:
        notes_upload = st.file_uploader("Latest meeting transcript or notes", type=SUPPORTED_NOTES, help="Supported: PDF, DOCX, TXT", key="notes_upload")

    demo_loaded = bool(st.session_state.get("demo_loaded"))
    requirements: DocumentContent | None = None
    notes: DocumentContent | None = None
    if requirements_upload is not None or demo_loaded:
        try:
            requirements = get_document_from_upload(requirements_upload, sample_document("requirements") if demo_loaded else None)
        except Exception as error:
            st.error(f"Could not parse {requirements_upload.name if requirements_upload else 'the requirements document'}: {error}")
    if notes_upload is not None or demo_loaded:
        try:
            notes = get_document_from_upload(notes_upload, sample_document("notes") if demo_loaded else None)
        except Exception as error:
            st.error(f"Could not parse {notes_upload.name if notes_upload else 'the meeting notes'}: {error}")

    if requirements and notes:
        st.markdown('<div class="section-label">2. Confirm extracted content</div>', unsafe_allow_html=True)
        preview_columns = st.columns(2)
        with preview_columns[0]:
            render_document_preview(requirements, "Requirements")
        with preview_columns[1]:
            render_document_preview(notes, "Meeting notes")

        st.markdown('<div class="section-label">3. Compare documents</div>', unsafe_allow_html=True)
        st.caption("The comparison looks for proposed changes, validation gaps, contradictions, repeated requirements, open questions, and explicitly stated actions.")
        if st.button("Compare documents", type="primary", use_container_width=True):
            with st.spinner("Reading both documents and preparing an evidence-backed comparison…"):
                ai_result, ai_error = run_ai_comparison(requirements, notes)
                if ai_result is not None:
                    comparison = ai_result
                    st.session_state["comparison_mode"] = "live"
                else:
                    comparison = local_compare(requirements, notes)
                    st.session_state["comparison_mode"] = "demo"
                st.session_state["comparison"] = comparison
                st.session_state["comparison_error"] = ai_error
                st.session_state["comparison_filenames"] = (requirements.filename, notes.filename)

    if st.session_state.get("comparison"):
        st.markdown('<div class="section-label">4. Review and approve</div>', unsafe_allow_html=True)
        if st.session_state.get("comparison_mode") == "live":
            st.success("Live OpenAI comparison complete. Review every finding before using it.")
        else:
            st.info("Transparent comparison mode is active because OPENAI_API_KEY is not configured. The app still used the extracted contents of the two documents; add the secret to enable the live model.")
        if st.session_state.get("comparison_error"):
            st.warning(st.session_state["comparison_error"])
        render_results(st.session_state["comparison"])


if __name__ == "__main__":
    main()