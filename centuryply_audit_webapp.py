# centuryply_audit_webapp.py
import os, re, math, uuid, io, datetime as dt
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, send_file, send_from_directory, flash
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from io import BytesIO

# Try pdf2image for thumbnails, but app works without it
try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False

# --- CONFIG ---
BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"
STATIC_DIR = BASE_DIR / "static"
THUMBS_DIR = STATIC_DIR / "thumbnails"
for d in (UPLOAD_DIR, REPORTS_DIR, STATIC_DIR, THUMBS_DIR):
    d.mkdir(parents=True, exist_ok=True)

ALLOWED_EXT = {'.xlsx', '.xlsm', '.xls'}

# Exact expected headers (user provided)
EXPECTED_HEADERS = [
    "Sl.", "Name", "Mobile Number", "Call Duration", "Call Date", "Audit Date", "Franchise",
    "Introduction", "Project Registration", "Product & Pricing requirement", "Product FeedBack",
    "Cross & Upsell of product", "Marketing Benefit", "Redemption", "Call Closure", "GTM Adherence",
    "CRM Update", "Softskill", "Total Score", "Total", "%age", "Audit Observation"
]
# Determine parameter columns automatically if present
PARAM_COLS = ["Introduction", "Project Registration", "Product & Pricing requirement", "Product FeedBack",
              "Cross & Upsell of product", "Marketing Benefit", "Redemption", "Call Closure",
              "GTM Adherence", "CRM Update", "Softskill"]

# --- FLASK APP ---
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"), static_folder=str(STATIC_DIR))
app.secret_key = os.environ.get("FLASK_SECRET", "replace-me-securely")

# --- Utilities ---
def allowed_file(filename):
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_EXT

def duration_to_minutes(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, pd.Timedelta):
        return x.total_seconds() / 60.0
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip()
    if s.lower() in ("no call", "no auditable call", "na", "nan", ""):
        return np.nan
    # common formats: HH:MM:SS or MM:SS or MM.SS or numeric minutes
    parts = re.split('[:.]', s)
    parts = [p for p in parts if p != ""]
    try:
        nums = [int(p) for p in parts]
    except:
        try:
            return float(s)
        except:
            return np.nan
    if len(nums) == 3:
        h, m, sec = nums
        return h * 60 + m + sec / 60.0
    if len(nums) == 2:
        m, sec = nums
        return m + sec / 60.0
    if len(nums) == 1:
        return float(nums[0])
    return np.nan

def safe_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).rstrip()

def fig_to_buf(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=180, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

# --- Routes ---
@app.route("/")
def index():
    logo_exists = (STATIC_DIR / "logo.png").exists() or (STATIC_DIR / "logo.jpg").exists()
    return render_template("upload.html", logo_exists=logo_exists)

@app.route("/static/<path:filename>")
def static_proxy(filename):
    return send_from_directory(str(STATIC_DIR), filename)

@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    logo = request.files.get("logo")
    report_name = request.form.get("report_name") or f"CenturyPly_Call_Audit_Report_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    if not f or f.filename == "":
        flash("Please upload an Excel file.", "danger")
        return redirect(url_for("index"))
    if not allowed_file(f.filename):
        flash("File type not allowed. Please upload an Excel file (.xlsx).", "danger")
        return redirect(url_for("index"))
    # persistent logo save
    if logo and logo.filename:
        logo_path = STATIC_DIR / "logo.png"
        logo.save(str(logo_path))
    # save upload
    fname = f"{uuid.uuid4().hex}_{safe_filename(f.filename)}"
    upath = UPLOAD_DIR / fname
    f.save(str(upath))
    # process and generate report
    try:
        pdf_path, xlsx_path = process_and_generate_reports(str(upath), report_name)
    except Exception as e:
        # show stack in logs and user-friendly message
        import traceback, sys
        traceback.print_exc()
        flash(f"Error generating report: {str(e)}", "danger")
        return redirect(url_for("index"))
    # generate thumbnail if possible
    try:
        if PDF2IMAGE_AVAILABLE:
            images = convert_from_bytes(open(pdf_path, "rb").read(), first_page=1, last_page=1, size=(400, None))
            if images:
                thumb_path = THUMBS_DIR / f"{pdf_path.stem}.png"
                images[0].save(str(thumb_path), "PNG")
    except Exception:
        pass
    return redirect(url_for("report_ready", filename=pdf_path.name))

@app.route("/report_ready/<filename>")
def report_ready(filename):
    return render_template("ready.html", filename=filename, report_url=url_for("download_report", filename=filename))

@app.route("/reports")
def reports():
    files = sorted(REPORTS_DIR.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    reports_list = []
    for p in files:
        thumb = THUMBS_DIR / f"{p.stem}.png"
        reports_list.append({
            "name": p.name,
            "pdf_url": url_for("download_report", filename=p.name),
            "thumb": url_for("static", filename=f"thumbnails/{thumb.name}") if thumb.exists() else None,
            "modified": dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        })
    return render_template("reports.html", reports=reports_list)

@app.route("/reports/<path:filename>")
def download_report(filename):
    p = REPORTS_DIR / filename
    if not p.exists():
        return "Not found", 404
    return send_file(str(p), as_attachment=True, download_name=p.name, mimetype="application/pdf")

@app.route("/reports/xlsx/<path:filename>")
def download_xlsx(filename):
    p = REPORTS_DIR / filename
    if not p.exists():
        return "Not found", 404
    return send_file(str(p), as_attachment=True, download_name=p.name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Core processing / report generation ---
def process_and_generate_reports(uploaded_excel_path, report_base_name):
    # Read Excel - try 'Followup' first, else first sheet
    xls = pd.ExcelFile(uploaded_excel_path)
    sheet = "Followup" if "Followup" in xls.sheet_names else xls.sheet_names[0]
    raw = xls.parse(sheet, header=None)
    # reconstruct headers if first two rows contain multi-line headers
    row0 = raw.iloc[0].fillna('')
    row1 = raw.iloc[1].fillna('')
    new_cols = []
    for a, b in zip(row0.tolist(), row1.tolist()):
        a = str(a).strip()
        b = str(b).strip()
        if a and a.lower() not in ('nan', ''):
            col = f"{a} - {b}" if b and b.lower() not in ('nan', '') else a
        else:
            col = b if b and b.lower() not in ('nan', '') else ''
        new_cols.append(col)
    df = raw.iloc[2:].copy().reset_index(drop=True)
    df.columns = new_cols
    # If new_cols missing expected names, fallback to original column names (Excel with single header)
    if not set(["Name", "Franchise", "%age"]).issubset(set(df.columns)):
        # try reading with first row as header
        df2 = pd.read_excel(uploaded_excel_path, sheet_name=sheet)
        df = df2.copy()
    # Standardize expected columns
    df.columns = [str(c).strip() for c in df.columns]
    # Ensure required columns exist
    for req in ("Name", "Franchise", "%age"):
        if req not in df.columns:
            raise ValueError(f"Required column '{req}' not found in uploaded sheet.")
    # Convert numeric cols
    for p in PARAM_COLS:
        if p in df.columns:
            df[p] = pd.to_numeric(df[p], errors='coerce')
    df["Total Score"] = pd.to_numeric(df.get("Total Score", pd.Series(dtype=float)), errors='coerce')
    df["%age"] = pd.to_numeric(df.get("%age", pd.Series(dtype=float)), errors='coerce')
    # call duration conversion
    if "Call Duration" in df.columns:
        df["Call_Duration_mins"] = df["Call Duration"].apply(duration_to_minutes)
    else:
        df["Call_Duration_mins"] = np.nan
    # Drop completely empty columns
    df = df.dropna(axis=1, how='all')
    # Focus on audited rows
    audited = df[df["%age"].notna()].copy()
    # Normalize param scaling if necessary
    for p in PARAM_COLS:
        if p in audited.columns:
            col = audited[p]
            if col.dropna().shape[0] > 0 and col.mean() > 1.0:
                if col.max() > 100:
                    audited[p] = col / 100.0
                else:
                    audited[p] = col / 10.0
    # Compute stats
    team_stats = audited.groupby("Franchise").agg(
        audits=("Name", "count"),
        avg_percentage=("%age", "mean"),
        median_percentage=("%age", "median"),
        std_percentage=("%age", "std"),
        avg_total_score=("Total Score", "mean"),
        avg_call_duration=("Call_Duration_mins", "mean")
    ).reset_index().sort_values("avg_percentage", ascending=False)
    rm_stats = audited.groupby(["Franchise", "Name"]).agg(
        audits=("Name", "count"),
        avg_percentage=("%age", "mean"),
        median_percentage=("%age", "median"),
        std_percentage=("%age", "std"),
        avg_total_score=("Total Score", "mean"),
        avg_call_duration=("Call_Duration_mins", "mean")
    ).reset_index().sort_values(["Franchise", "avg_percentage"], ascending=[True, False])

    team_param = audited.groupby("Franchise")[[c for c in PARAM_COLS if c in audited.columns]].mean().reset_index()
    rm_param = audited.groupby(["Franchise", "Name"])[[c for c in PARAM_COLS if c in audited.columns]].mean().reset_index()

    # Generate Excel summary
    safe_base = safe_filename(report_base_name)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_name = f"{safe_base}_{timestamp}.xlsx"
    excel_path = REPORTS_DIR / excel_name
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        team_stats.to_excel(writer, sheet_name="Team_Summary", index=False)
        rm_stats.to_excel(writer, sheet_name="RM_Summary", index=False)
        if not team_param.empty:
            team_param.to_excel(writer, sheet_name="Team_Params", index=False)
        audited.to_excel(writer, sheet_name="Audited_Rows", index=False)

    # Generate PDF report (multi-page) - follow required sections
    pdf_name = f"{safe_base}_{timestamp}.pdf"
    pdf_path = REPORTS_DIR / pdf_name
    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Title", fontSize=20, leading=22, alignment=1, textColor=colors.HexColor("#C00000"), spaceAfter=12, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="SubHeader", fontSize=14, leading=16, textColor=colors.HexColor("#C00000"), spaceAfter=6, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="NormalSmall", fontSize=9, leading=11))
    elements = []
    # Header/cover
    elements.append(Paragraph("CenturyPly Call Audit Report", styles["Title"]))
    elements.append(Paragraph(f"Consolidated Period: Generated on {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["NormalSmall"]))
    elements.append(Spacer(1, 12))
    elements.append(PageBreak())
    # Table of Contents (simple)
    elements.append(Paragraph("Table of Contents", styles["SubHeader"]))
    toc = []
    toc.append("1. Cover Page")
    toc.append("2. Table of Contents")
    toc.append("3. Team-wise summary with bar charts")
    toc.append("4. RM-wise performance table")
    toc.append("5. Parameter-wise average table")
    toc.append("6. Call Duration vs % correlation trend line")
    toc.append("7. Top 5 and Bottom 5 RMs (with mini bar charts)")
    toc.append("8. Rest of the RMs (with mini bar charts)")
    toc.append("9. Findings & Recommendations for each RM and team")
    toc.append("10. Appendix — data overview")
    elements.append(Paragraph("<br/>".join([f"• {t}" for t in toc]), styles["NormalSmall"]))
    elements.append(PageBreak())

    # Per-team sections
    for _, trow in team_stats.iterrows():
        team = trow["Franchise"]
        elements.append(Paragraph(f"Team: {team}", styles["SubHeader"]))
        # summary table
        sum_data = [["Metric", "Value"],
                    ["Total Audits", f"{int(trow['audits'])}"],
                    ["Average %", f"{trow['avg_percentage']*100:.2f}%"],
                    ["Median %", f"{trow['median_percentage']*100:.2f}%"],
                    ["Std Dev", f"{trow['std_percentage']*100:.2f}%"],
                    ["Avg Call Duration (mins)", f"{trow['avg_call_duration']:.2f}"]]
        t = Table(sum_data, colWidths=[220, 140])
        t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4B4B4B")), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.25,colors.lightgrey)]))
        elements.append(t); elements.append(Spacer(1,8))
        # parameter chart
        if team_param.shape[0] > 0 and any(col in team_param.columns for col in PARAM_COLS):
            rowp = team_param[team_param['Franchise']==team]
            if not rowp.empty:
                vals = rowp[[c for c in PARAM_COLS if c in rowp.columns]].iloc[0].values * 100
                params = [c for c in PARAM_COLS if c in rowp.columns]
                fig, ax = plt.subplots(figsize=(8,2.4))
                ax.bar(params, vals, color="#C00000", edgecolor="#4B4B4B")
                ax.set_ylim(0,100)
                ax.tick_params(axis='x', rotation=45, labelsize=8)
                plt.grid(axis='y', linestyle='--', alpha=0.4)
                elements.append(Image(fig_to_buf(fig), width=520, height=140))
                elements.append(Spacer(1,6))
                # parameter table
                pdata = [["Parameter", "Average %"]]
                for p,v in zip(params, vals):
                    pdata.append([p, f"{v:.1f}%"])
                pt = Table(pdata, colWidths=[300,80])
                pt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#C00000")),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.lightgrey)]))
                elements.append(pt)
        # call duration vs % trend for this team
        team_rm = rm_stats[rm_stats["Franchise"]==team].copy()
        # merge avg duration
        avg_dur = audited.groupby(['Franchise','Name'])['Call_Duration_mins'].mean().reset_index().rename(columns={'Call_Duration_mins':'Avg_Duration_mins'})
        team_rm = team_rm.merge(avg_dur[avg_dur['Franchise']==team][['Name','Avg_Duration_mins']], on='Name', how='left')
        x = team_rm['Avg_Duration_mins'].fillna(0).values
        y = team_rm['avg_percentage'].fillna(0).values * 100
        if len(x)>0:
            fig, ax = plt.subplots(figsize=(7,2.6))
            ax.scatter(x, y, color="#C00000", alpha=0.8)
            if len(x)>1 and np.nanstd(x)>0:
                coeffs = np.polyfit(np.nan_to_num(x), np.nan_to_num(y), 1)
                poly = np.poly1d(coeffs)
                xs = np.linspace(np.nanmin(x), np.nanmax(x), 100)
                ax.plot(xs, poly(xs), color="#4B4B4B", linestyle='--', linewidth=1)
                corr = np.corrcoef(np.nan_to_num(x), np.nan_to_num(y))[0,1]
            else:
                corr = 0.0
            ax.set_xlabel("Avg Call Duration (mins)"); ax.set_ylabel("Avg %age")
            ax.set_title("Call Duration vs. Audit % (RM level) — Trend Line")
            plt.grid(axis='y', linestyle='--', alpha=0.4)
            elements.append(Image(fig_to_buf(fig), width=420, height=160))
            elements.append(Paragraph(f"Correlation (duration vs %): {corr:.2f}", styles["NormalSmall"]))
        # top 5 and bottom 5
        top5 = team_rm.sort_values('avg_percentage', ascending=False).head(5)
        bot5 = team_rm.sort_values('avg_percentage', ascending=True).head(5)
        def small_table(df, title):
            data = [[title, 'Avg %', 'Audits']]
            for _, r in df.iterrows():
                data.append([r['Name'], f"{r['avg_percentage']*100:.1f}%", int(r['audits'])])
            return Table(data, colWidths=[240,80,80])
        elements.append(Spacer(1,6))
        elements.append(Paragraph("Top 5 Performing RMs", styles['SubHeader']))
        elements.append(small_table(top5, 'Top 5'))
        # mini-chart top5
        if not top5.empty:
            fig, ax = plt.subplots(figsize=(4,1.6))
            vals = (top5['avg_percentage']*100).tolist()
            names = top5['Name'].tolist()
            ax.barh(range(len(vals)), vals, color="#C00000")
            ax.set_yticks(range(len(vals))); ax.set_yticklabels(names, fontsize=7); ax.invert_yaxis(); ax.set_xlim(0,100)
            elements.append(Image(fig_to_buf(fig), width=350, height=90))
        elements.append(Spacer(1,6))
        elements.append(Paragraph("Bottom 5 Performing RMs", styles['SubHeader']))
        elements.append(small_table(bot5, 'Bottom 5'))
        if not bot5.empty:
            fig, ax = plt.subplots(figsize=(4,1.6))
            vals = (bot5['avg_percentage']*100).tolist()
            names = bot5['Name'].tolist()
            ax.barh(range(len(vals)), vals, color="#4B4B4B")
            ax.set_yticks(range(len(vals))); ax.set_yticklabels(names, fontsize=7); ax.invert_yaxis(); ax.set_xlim(0,100)
            elements.append(Image(fig_to_buf(fig), width=350, height=90))
        elements.append(PageBreak())

    # RM-level: Rest of RMs & recommendations
    elements.append(Paragraph("RM-Level Findings & Recommendations (All RMs)", styles['SubHeader']))
    all_rms = rm_param.copy().merge(rm_stats[['Franchise','Name','audits','avg_percentage','avg_total_score']], on=['Franchise','Name'], how='left')
    all_rms = all_rms.sort_values(['Franchise','avg_percentage'], ascending=[True,False])
    count = 0
    for idx, row in all_rms.iterrows():
        name = row['Name']; team = row['Franchise']
        avg_pct = (row['avg_percentage']*100) if not pd.isna(row['avg_percentage']) else 0.0
        audits_count = int(row['audits']) if not pd.isna(row['audits']) else 0
        params = row[[c for c in PARAM_COLS if c in row.index]]
        params_sorted = params.sort_values(ascending=False)
        strengths = ", ".join([f"{p} ({params_sorted[p]*100:.0f}%)" for p in params_sorted.head(3).index])
        weaknesses = ", ".join([f"{p} ({params_sorted[p]*100:.0f}%)" for p in params_sorted.tail(3).index])
        recs = f"• Improve {params_sorted.tail(1).index[0]} through role-play and checklist reviews. • Review top-performer calls weekly. • Use CRM reminders to improve follow-up adherence."
        block = f"<b>{name} — {team}</b><br/>Avg %: {avg_pct:.1f}%, Audits: {audits_count}<br/><b>Strengths:</b> {strengths}<br/><b>Improvements:</b> {weaknesses}<br/>{recs}"
        elements.append(Paragraph(block, styles['NormalSmall']))
        elements.append(Spacer(1,6))
        count += 1
        if count % 22 == 0:
            elements.append(PageBreak())

    # Consolidated overview & appendix
    elements.append(PageBreak())
    elements.append(Paragraph("Consolidated Overview & Company Recommendations", styles['SubHeader']))
    fig, ax = plt.subplots(figsize=(8,2.4))
    teams = team_stats['Franchise']
    vals = team_stats['avg_percentage']*100
    ax.bar(teams, vals, color="#C00000", edgecolor="#4B4B4B")
    ax.set_ylim(0,100); ax.tick_params(axis='x', rotation=45, labelsize=8)
    plt.grid(axis='y', linestyle='--', alpha=0.4)
    elements.append(Image(fig_to_buf(fig), width=520, height=140))
    elements.append(Spacer(1,8))
    overall_param_avg = aud[[c for c in PARAM_COLS if c in aud.columns]].mean().sort_values()
    if not overall_param_avg.empty:
        overall_top = overall_param_avg.tail(3); overall_bot = overall_param_avg.head(3)
        summary_notes = "<b>Top Parameters (Company-wide):</b><br/>" + "<br/>".join([f"• {p}: {float(v*100):.1f}%" for p,v in overall_top.items()])
        summary_notes += "<br/><br/><b>Areas to Improve (Company-wide):</b><br/>" + "<br/>".join([f"• {p}: {float(v*100):.1f}%" for p,v in overall_bot.items()])
        summary_notes += "<br/><br/><b>Company Recommendations:</b><br/>• Implement monthly parameter-wise review meetings.<br/>• Use mentor pairing to lift bottom-5 RMs.<br/>• Incorporate CRM update checks into audit feedback.<br/>• Track progress via weekly dashboard."
        elements.append(Paragraph(summary_notes, styles['NormalSmall']))

    # Appendix - data overview
    elements.append(PageBreak())
    elements.append(Paragraph("Appendix — Data Overview (Top rows)", styles['SubHeader']))
    sample_table = [["Column", "Sample Values"]]
    for c in df.columns[:10]:
        vals = df[c].dropna().astype(str).unique()[:3]
        sample_table.append([c, ", ".join(vals)])
    elements.append(Table(sample_table, colWidths=[220, 420]))
    doc.build(elements)

    # return paths
    return pdf_path, excel_path

if __name__ == "__main__":
    # for local testing, enable debug
    app.run(host="0.0.0.0", port=5000, debug=True)
