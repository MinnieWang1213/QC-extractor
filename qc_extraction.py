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

            # Split CATEGORY and SYMBOL
            if "；" in category_raw:
                category_main, symbol = [part.strip() for part in category_raw.split("；", 1)]
            else:
                category_main, symbol = category_raw.strip(), ""

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

        # Create DataFrame and set column order
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

        # Offer download
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name="Evaluation_Summary_with_CATEGORY_and_ERROR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
