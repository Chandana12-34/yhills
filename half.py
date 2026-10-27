if CACHE_FILE.exists():

        try:

            CACHE = json.loads(CACHE_FILE.read_text(encoding="utf-8"))

        except (OSError, json.JSONDecodeError):

            logger.warning("Could not load cache; starting with an empty cache.")

            CACHE = {}

def save_cache():

    # Store only successful results.

    temp_file = CACHE_FILE.with_suffix(".tmp")

    temp_file.write_text(

        json.dumps(CACHE, ensure_ascii=False, indent=2),

        encoding="utf-8",

    )

    temp_file.replace(CACHE_FILE)

def analyze_ticket(ticket):

    base = {column: "" for column in OUTPUT_COLUMNS}

    for field in [

        "ritm",

        "state",

        "priority",

        "short_description",

        "assignment_group",

        "assigned_to",

        "created",

    ]:

        base[field] = ticket.get(field, "")

    base["analysis_status"] = "Failed"

    base["analysis_error"] = ""

    if not ticket["combined_text"].strip():

        base["user_question"] = "Question not clearly stated in the ticket."

        base["answer_or_solution"] = (

            "No answer or solution documented in the ticket."

        )

        base["analysis_status"] = "Skipped"

        base["analysis_error"] = "No usable request, comments, or resolution text."

        return base

    cache_key = ticket["ticket_hash"]

    if cache_key in CACHE:

        base.update(CACHE[cache_key])

        base["analysis_status"] = "Success"

        base["analysis_error"] = ""

        return base

    try:

        response = call_llm(build_prompt(ticket))

        analysis = parse_llm_json(response)

        base.update(analysis)

        base["analysis_status"] = "Success"

        CACHE[cache_key] = analysis

    except Exception as exc:

        logger.error(

            "Failed to analyze RITM %s: %s",

            ticket.get("ritm", "Unknown"),

            exc,

        )

        base["analysis_error"] = str(exc)

    return base

# ============================================================================

# REPORT SUMMARIES

# ============================================================================

def build_overview(results):

    successful = results[results["analysis_status"] == "Success"]

    no_answer = successful["answer_or_solution"].str.contains(

        r"No answer or solution documented",

        case=False,

        na=False,

    )

    return pd.DataFrame(

        [

            ("Total RITMs selected", len(results)),

            ("Successfully analyzed", len(successful)),

            ("Failed analyses", int((results["analysis_status"] == "Failed").sum())),

            ("Skipped tickets", int((results["analysis_status"] == "Skipped").sum())),

            ("Tickets without documented answers", int(no_answer.sum())),

            (

                "Tickets marked as needing training",

                int((successful["training_needed"] == "Yes").sum()),

            ),

            (

                "High-confidence analyses",

                int((successful["confidence"] == "High").sum()),

            ),

            (

                "Unique RITMs",

                int(results["ritm"].nunique()),

            ),

        ],

        columns=["Metric", "Value"],

    )

def build_group_summary(df, column):

    successful = df[df["analysis_status"] == "Success"].copy()

    if successful.empty:

        return pd.DataFrame(columns=[column, "Ticket Count"])

    summary = (

        successful.groupby(column, dropna=False)

        .size()

        .reset_index(name="Ticket Count")

        .sort_values("Ticket Count", ascending=False)

    )

    return summary

def build_recurring_questions(df):

    successful = df[df["analysis_status"] == "Success"].copy()

    if successful.empty:

        return pd.DataFrame(

            columns=[

                "User Question",

                "Ticket Count",

                "Example RITMs",

                "Documented Answers",

                "Training Topic",

            ]

        )

    # Exact normalized wording groups identical questions.

    # Similar but differently worded questions should be reviewed together.

    successful["question_key"] = (

        successful["user_question"]

        .fillna("")

        .str.lower()

        .str.replace(r"[^a-z0-9\s]", "", regex=True)

        .str.replace(r"\s+", " ", regex=True)

        .str.strip()

    )

    rows = []

    for _, group in successful.groupby("question_key", dropna=False):

        rows.append(

            {

                "User Question": group["user_question"].iloc[0],

                "Ticket Count": len(group),

                "Example RITMs": ", ".join(

                    group["ritm"].astype(str).head(10)

                ),

                "Documented Answers": " | ".join(

                    dict.fromkeys(

                        value for value in group["answer_or_solution"].astype(str)

                        if value.strip()

                    )

                )[:3000],

                "Training Topic": "; ".join(

                    dict.fromkeys(

                        value for value in group["training_topic"].astype(str)

                        if value.strip() and value != "Not identified"

                    )

                )[:1000],

            }

        )

    return pd.DataFrame(rows).sort_values(

        "Ticket Count", ascending=False

    )

def build_category_breakdown(df):

    successful = df[df["analysis_status"] == "Success"]

    if successful.empty:

        return pd.DataFrame(

            columns=["Issue Type", "Category", "Subcategory", "Ticket Count"]

        )

    return (

        successful.groupby(

            ["issue_type", "category", "subcategory"],

            dropna=False,

        )

        .size()

        .reset_index(name="Ticket Count")

        .rename(

            columns={

                "issue_type": "Issue Type",

                "category": "Category",

                "subcategory": "Subcategory",

            }

        )

        .sort_values("Ticket Count", ascending=False)

    )

# ============================================================================

# EXCEL DASHBOARD AND CHARTS

# ============================================================================

def write_report(results, output_file):

    overview = build_overview(results)

    issue_summary = build_group_summary(results, "issue_type")

    training_summary = build_group_summary(results, "training_needed")

    recurring_questions = build_recurring_questions(results)

    category_breakdown = build_category_breakdown(results)

    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        results.to_excel(writer, sheet_name="RITM Analysis", index=False)

        overview.to_excel(writer, sheet_name="Dashboard", index=False)

        issue_summary.to_excel(writer, sheet_name="Issue Breakdown", index=False)

        recurring_questions.to_excel(

            writer, sheet_name="Recurring Questions", index=False

        )

        category_breakdown.to_excel(

            writer, sheet_name="Category Breakdown", index=False

        )

        training_summary.to_excel(

            writer, sheet_name="Training Summary", index=False

        )

        wb = writer.book

        # Consistent report formatting.

        for ws in wb.worksheets:

            ws.freeze_panes = "A2"

            ws.auto_filter.ref = ws.dimensions

            ws.sheet_view.showGridLines = False

            for cell in ws[1]:

                cell.font = Font(bold=True, color="FFFFFF")

                cell.fill = PatternFill(

                    fill_type="solid",

                    fgColor="17365D",

                )

                cell.alignment = Alignment(

                    horizontal="center",

                    vertical="center",

                    wrap_text=True,

                )

            ws.row_dimensions[1].height = 32

            for column_cells in ws.columns:

                letter = column_cells[0].column_letter

                max_length = max(

                    (

                        len(str(cell.value))

                        for cell in column_cells[:500]

                        if cell.value is not None

                    ),

                    default=10,

                )

                ws.column_dimensions[letter].width = min(

                    max(max_length + 2, 12), 48

                )

            for row in ws.iter_rows(min_row=2):

                for cell in row:

                    cell.alignment = Alignment(

                        vertical="top",

                        wrap_text=True,

                    )

        # Dashboard title.

        dashboard = wb["Dashboard"]

        dashboard.insert_rows(1, 2)

        dashboard["A1"] = "MAOS KNOWLEDGE GAP ANALYSIS"

        dashboard["A1"].font = Font(

            size=18,

            bold=True,

            color="17365D",

        )

        dashboard["A2"] = "ServiceNow RITM analysis and knowledge improvement"

        dashboard["A2"].font = Font(italic=True, color="666666")

        dashboard.freeze_panes = "A4"

        # Issue type chart.

        issue_ws = wb["Issue Breakdown"]

        if issue_ws.max_row > 1:

            chart = BarChart()

            chart.type = "bar"

            chart.style = 10

            chart.title = "Tickets by Issue Type"

            chart.y_axis.title = "Issue Type"

            chart.x_axis.title = "Ticket Count"

            chart.height = 7

            chart.width = 14

            chart.add_data(

                Reference(

                    issue_ws,

                    min_col=2,

                    min_row=1,

                    max_row=issue_ws.max_row,

                ),

                titles_from_data=True,

            )

            chart.set_categories(

                Reference(

                    issue_ws,

                    min_col=1,

                    min_row=2,

                    max_row=issue_ws.max_row,

                )

            )

            dashboard.add_chart(chart, "D4")

        # Training chart.

        training_ws = wb["Training Summary"]

        if training_ws.max_row > 1:

            pie = PieChart()

            pie.title = "Training Requirement"

            pie.height = 7

            pie.width = 11

            pie.add_data(

                Reference(

                    training_ws,

                    min_col=2,

                    min_row=1,

                    max_row=training_ws.max_row,

                ),

                titles_from_data=True,

            )

            pie.set_categories(

                Reference(

                    training_ws,

                    min_col=1,

                    min_row=2,

                    max_row=training_ws.max_row,

                )

            )

            pie.dataLabels = DataLabelList()

            pie.dataLabels.showPercent = True

            pie.dataLabels.showLeaderLines = True

            dashboard.add_chart(pie, "D19")

        # Make the detailed analysis easier to read.

        analysis_ws = wb["RITM Analysis"]

        analysis_ws.freeze_panes = "A2"

    logger.info("Excel report written to: %s", output_file.resolve())

# ============================================================================

# MAIN

# ============================================================================

def main():

    logger.info("Starting MAOS Knowledge Gap Analyzer")

    df, mapping = load_excel()

    tickets = [

        preprocess_ticket(row, mapping)

        for _, row in df.iterrows()

    ]

    # Keep the first 100 rows only if the workbook is already sorted with

    # the intended most recent/top tickets first.

    # Otherwise sort/filter the export correctly before running this script.

    if MAX_TICKETS > 0:

        tickets = tickets[:MAX_TICKETS]

    logger.info("Selected %d RITMs for analysis", len(tickets))

    load_cache()

    results = [None] * len(tickets)

    # Concurrent requests improve throughput.

    # Reduce MAX_WORKERS if your internal API has rate limits.

    with ThreadPoolExecutor(max_workers=max(1, MAX_WORKERS)) as executor:

        future_map = {

            executor.submit(analyze_ticket, ticket): index

            for index, ticket in enumerate(tickets)

        }

        completed = 0

        for future in as_completed(future_map):

            index = future_map[future]

            ticket = tickets[index]

            try:

                results[index] = future.result()

            except Exception as exc:

                logger.exception("Unexpected worker error for %s", ticket["ritm"])

                results[index] = {

                    column: ticket.get(column, "")

                    for column in OUTPUT_COLUMNS

                }

                results[index]["analysis_status"] = "Failed"

                results[index]["analysis_error"] = str(exc)

            completed += 1

            logger.info(

                "Progress %d/%d | RITM %s | %s",

                completed,

                len(tickets),

                ticket["ritm"],

                results[index]["analysis_status"],

            )

            # Save successful cache entries periodically.

            if completed % 10 == 0:

                save_cache()

    save_cache()

    results_df = pd.DataFrame(results, columns=OUTPUT_COLUMNS)

    write_report(results_df, OUTPUT_FILE)

    success_count = int((results_df["analysis_status"] == "Success").sum())

    failed_count = int((results_df["analysis_status"] == "Failed").sum())

    skipped_count = int((results_df["analysis_status"] == "Skipped").sum())

    logger.info("=" * 65)

    logger.info("MAOS ANALYSIS COMPLETED")

    logger.info("Selected: %d", len(results_df))

    logger.info("Successful: %d", success_count)

    logger.info("Failed: %d", failed_count)

    logger.info("Skipped: %d", skipped_count)

    logger.info("Report: %s", OUTPUT_FILE.resolve())

    logger.info("=" * 65)

if __name__ == "__main__":

    main()