# -*- coding: utf-8 -*-
"""QC Extraction with CATEGORY normalization and ERROR field (preserve unmatched)"""

import streamlit as st
from docx import Document
import docx2txt
import pandas as pd
import re
import io

st.title("📝 Evaluation Form Extractor")
st.write("Upload one or more `.docx` evaluation forms to extract structured feedback and ratings.")

uploaded_files = st.file_uploader("Upload Word Files", type="docx", accept_multiple_files=True)

def extract_field(text, label):
    pattern = rf"{label}\s*\n([^\n]+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def extract_rating(text):
    rating_block = re.findall(r"Quality rating[\s\S]{0,150}", text, re.IGNORECASE)
    for block in rating_block:
        match = re.search(r"\b([1-6])\b", block)
        if match:
            return match.group(1)
    return ""

def check_alert(text):
    feedback_pattern = r"Evaluator’s feedback \(including any guidance\) on the assignment[\s\S]{0,500}"
    match = re.search(feedback_pattern, text)
    if match:
        feedback = match.group()
        if any(term in feedback for term in ["温馨提醒：", "温馨提示：", "提醒：", "提示："]):
            return "ALERT"
    return ""

error_types = [
    "Misformatted texts/tables/charts (not impacting content)",
    "Missing punctuation (not impacting content)",
    "Wrong space before/after titles",
    "Missing asterisks, parentheses, square brackets, capital/lower-case letter, recycle, end line",
    "Misformatted text/tables/charts in budget documents",
    "Any mistake on the cover page including TOC",
    "Missing punctuation (impact content)",
    "Editor/Translator’s corrections/mark-ups missed",
    "Any substantive change in text without Editor/Translator's approval",
    "Obvious language mistakes",
    "Missing/wrong/misplaced figures, dates, names, countries, symbols",
    "Content differs from submitted translation",
    "Wrong numbering",
    "Wrong document type",
    "Wrong/missing headers/footers",
    "Hyperlinks not activated or used as required",
    "Document not uploaded on gDoc or labelled properly",
    "Track changes not accepted and/or comments/mark-ups not removed"
]

VALID_CATEGORIES = [
    'Plain Text', 'Text with Tables/Charts', 'Text with many footnotes', 'Budget',
    'BoA', 'Letters/Reports', 'RES-related', 'Supplements', 'Publications', 'Other'
]

def normalize_category(raw_category):
    raw_category_clean = raw_category.strip()
    raw_lower = raw_category_clean.lower()
    for valid in VALID_CATEGORIES:
        if raw_lower == valid.lower():
            return valid  # Return standardized form
    return raw_category_clean  # Keep original if unmatched

def extract_errors_and_comments(doc):
    results = []
    shorthand_map = {'c': 'Content', 's': 'Style', 'p': 'Process'}
    for table in doc.tables:
        if len(table.rows[0].cells) >= 4 and table.rows[0].cells[3].text.strip().lower() == "issue":
            for row in table.rows[1:]:
                issue_text = row.cells[3].text.strip().lower()
                comment_text = row.cells[2].text.strip()
                first_line = comment_text.split("\n")[0].strip()
                matched_type = next((et for et in error_types if et.lower() in first_line.lower()), first_line)
                match = re.search(r'\b(content|style|process|c|s|p)\b(?:\s+(major|medium|minor))?', issue_text, re.IGNORECASE)
                if match:
                    raw_category, severity = match.groups()
                    category = shorthand_map.get(raw_category.lower(), raw_category.capitalize())
                    severity = severity if severity else "Unspecified"
                    results.append((category, severity, matched_type))
    return results

if uploaded_files:
    all_rows = []
    for uploaded_file in uploaded_files:
        raw_text = docx2txt.process(uploaded_file)
        doc = Document(uploaded_file)

        name = extract_field(raw_text, "Evaluatee")
        job = extract_field(raw_text, "Job No.")
        category_raw = extract_field(raw_text, "Category")
        assignment_type = extract_field(raw_text, "Assignment type")
        level = extract_field(raw_text, "Level of complexity")
        time_eff = extract_field(raw_text, "Time efficiency")
        rating = extract_rating(raw_text)
        alert_flag = check_alert(raw_text)

        # Split and normalize CATEGORY and SYMBOL
        if "；" in category_raw:
            raw_category_part, symbol = [part.strip() for part in category_raw.split("；", 1)]
        else:
            raw_category_part, symbol = category_raw.strip(), ""
        category_main = normalize_category(raw_category_part)

        errors = extract_errors_and_comments(doc)

        if errors:
            error_flag = "YES"
            for category, severity, type_of_error in errors:
                all_rows.append({
                    "NAME": name,
                    "JOB": job,
                    "CATEGORY": category_main,
                    "SYMBOL": symbol,
                    "TYPE OF ERROR": type_of_error,
                    "CATEGORY OF ERROR": category,
                    "SEVERITY OF ERROR": severity,
                    "ASSIGNMENT TYPE": assignment_type,
                    "LEVEL OF COMPLEXITY": level,
                    "TIME EFFICIENCY": time_eff,
                    "OVERALL RATING": rating,
                    "STATUS OF REPORT": "TO STAFF",
                    "ALERT": alert_flag,
                    "ERROR": error_flag
                })
        else:
            error_flag = "NO ERROR BUT ALERT" if alert_flag else "NO"
            all_rows.append({
                "NAME": name,
                "JOB": job,
                "CATEGORY": category_main,
                "SYMBOL": symbol,
                "TYPE OF ERROR": "",
                "CATEGORY OF ERROR": "",
                "SEVERITY OF ERROR": "",
                "ASSIGNMENT TYPE": assignment_type,
                "LEVEL OF COMPLEXITY": level,
                "TIME EFFICIENCY": time_eff,
                "OVERALL RATING": rating,
                "STATUS OF REPORT": "TO STAFF",
                "ALERT": alert_flag,
                "ERROR": error_flag
            })

    df = pd.DataFrame(all_rows)
    column_order = [
        "NAME", "JOB", "CATEGORY", "SYMBOL",
        "TYPE OF ERROR", "CATEGORY OF ERROR", "SEVERITY OF ERROR",
        "ASSIGNMENT TYPE", "LEVEL OF COMPLEXITY", "TIME EFFICIENCY", "OVERALL RATING",
        "STATUS OF REPORT", "ALERT", "ERROR"
    ]
    df = df[column_order]

    st.success("✅ Extraction complete!")
    st.dataframe(df)

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    st.download_button(
        label="📥 Download Excel File",
        data=output,
        file_name="Evaluation_Summary_with_CATEGORY_and_ERROR.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
