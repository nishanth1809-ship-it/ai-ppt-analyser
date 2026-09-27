import os
import re
import html
from flask import Flask, request, render_template_string, send_file
from pptx import Presentation
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib import colors

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

INDEX_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PPT Analyzer & Report Generator</title>

<style>
:root{
--bg:#08111f;--surface:#0f1f33;--surface2:#132942;--text:#eef7ff;--muted:#9bb0c7;
--primary:#22d3ee;--secondary:#3b82f6;--accent:#a78bfa;--border:#23415f;
--success:#34d399;--danger:#fb7185;--shadow:0 18px 50px #00000055;
}
*{box-sizing:border-box}
body{margin:0;font-family:Inter,Arial,sans-serif;background:radial-gradient(circle at top right,#182d55 0,#08111f 42%,#050b14 100%);color:var(--text);min-height:100vh}
.container{max-width:1120px;margin:auto;padding:35px 20px}
.hero{text-align:center;padding:28px 20px}
.badge{display:inline-block;padding:8px 14px;border:1px solid var(--border);border-radius:999px;color:var(--primary);background:#ffffff08;font-size:12px;font-weight:800;letter-spacing:1.5px}
.hero h1{font-size:46px;margin:18px 0 10px;background:linear-gradient(90deg,var(--primary),var(--accent),var(--secondary));-webkit-background-clip:text;background-clip:text;color:transparent}
.hero p{color:var(--muted);line-height:1.7;max-width:780px;margin:auto}
.themebar{display:flex;justify-content:center;gap:8px;flex-wrap:wrap;margin:18px 0}
.themebar button{border:1px solid var(--border);background:var(--surface);color:var(--text);padding:8px 13px;border-radius:999px;font-size:12px;cursor:pointer}
.themebar button:hover{border-color:var(--primary);transform:translateY(-1px)}
.box{background:linear-gradient(145deg,#ffffff0b,#ffffff04);border:1px solid var(--border);border-radius:24px;padding:42px;box-shadow:var(--shadow);text-align:center;margin-top:25px;backdrop-filter:blur(12px)}
.box h2{font-size:28px;margin:0 0 8px}
input[type=file]{margin:22px;padding:16px;border:1px dashed var(--primary);border-radius:14px;width:min(560px,90%);background:#071321;color:var(--muted)}
button,.download{background:linear-gradient(90deg,var(--secondary),var(--primary));color:#04111c;border:0;border-radius:12px;padding:14px 25px;font-size:16px;font-weight:800;cursor:pointer;text-decoration:none;display:inline-block;box-shadow:0 8px 25px #00000030}
button:hover,.download:hover{filter:brightness(1.08);transform:translateY(-1px)}
.features{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-top:22px}
.card{background:var(--surface);border:1px solid var(--border);padding:24px;border-radius:18px;text-align:center;transition:.2s}
.card:hover{transform:translateY(-4px);border-color:var(--primary)}
.card h3{color:var(--primary)}.card p{color:var(--muted);line-height:1.6}
.muted{color:var(--muted)}.actions{text-align:center;margin:30px 0}
@media(max-width:800px){.features{grid-template-columns:1fr 1fr}}
@media(max-width:500px){.features{grid-template-columns:1fr}.hero h1{font-size:34px}.box{padding:25px}}
</style>
<script>
function setTheme(name){
 const themes={
 ocean:{bg:"#08111f",surface:"#0f1f33",surface2:"#132942",text:"#eef7ff",muted:"#9bb0c7",primary:"#22d3ee",secondary:"#3b82f6",accent:"#a78bfa",border:"#23415f",success:"#34d399",danger:"#fb7185"},
 sunset:{bg:"#1b0c12",surface:"#2b141d",surface2:"#3a1b25",text:"#fff5ed",muted:"#d6aaa0",primary:"#fb7185",secondary:"#f97316",accent:"#facc15",border:"#63303a",success:"#4ade80",danger:"#ef4444"},
 emerald:{bg:"#071713",surface:"#0d241e",surface2:"#123329",text:"#effff7",muted:"#9cc8b5",primary:"#34d399",secondary:"#14b8a6",accent:"#a3e635",border:"#1d5545",success:"#4ade80",danger:"#fb7185"},
 royal:{bg:"#100b25",surface:"#1b1439",surface2:"#251b4e",text:"#f7f3ff",muted:"#b9acd8",primary:"#c084fc",secondary:"#6366f1",accent:"#f472b6",border:"#493b78",success:"#34d399",danger:"#fb7185"}
 };
 const t=themes[name]||themes.ocean;
 Object.entries(t).forEach(([k,v])=>document.documentElement.style.setProperty("--"+k,v));
 localStorage.setItem("pptTheme",name);
}
window.addEventListener("DOMContentLoaded",()=>setTheme(localStorage.getItem("pptTheme")||"ocean"));
</script>

</head>
<body>
<div class="container">
<div class="hero">
<div class="badge">AI PRESENTATION ANALYZER</div>
<h1>AI-Based PPT Analyzer</h1>
<p>Upload a PowerPoint presentation to analyze every slide, calculate scores, identify issues, generate paragraph suggestions, and download a final PDF report.</p>
<div class="themebar"><button type="button" onclick="setTheme('ocean')">Ocean</button><button type="button" onclick="setTheme('sunset')">Sunset</button><button type="button" onclick="setTheme('emerald')">Emerald</button><button type="button" onclick="setTheme('royal')">Royal Purple</button></div>
</div>
<div class="box">
<h2>Upload Your PPT / PPTX</h2>
<form action="/analyze" method="post" enctype="multipart/form-data">
<input type="file" name="ppt_file" accept=".ppt,.pptx" required><br>
<button type="submit">Analyze PPT</button>
</form>
<p class="muted">The analysis is based on slide text, titles, structure, content length and common presentation-quality checks.</p>
</div>
<div class="features">
<div class="card"><h3>Slide Analysis</h3><p>Extracts and checks every individual slide.</p></div>
<div class="card"><h3>Score</h3><p>Generates an overall and slide-wise score.</p></div>
<div class="card"><h3>Issues</h3><p>Explains detected issues for each slide.</p></div>
<div class="card"><h3>PDF Report</h3><p>Creates a downloadable paragraph-based report.</p></div>
</div>
</div>
</body>
</html>
"""

RESULT_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PPT Analysis Result</title>

<style>
:root{
--bg:#08111f;--surface:#0f1f33;--surface2:#132942;--text:#eef7ff;--muted:#9bb0c7;
--primary:#22d3ee;--secondary:#3b82f6;--accent:#a78bfa;--border:#23415f;
--success:#34d399;--danger:#fb7185;
}
*{box-sizing:border-box}
body{margin:0;font-family:Inter,Arial,sans-serif;background:radial-gradient(circle at top right,#182d55 0,#08111f 42%,#050b14 100%);color:var(--text);min-height:100vh}
.container{max-width:1120px;margin:auto;padding:30px 20px}
.top{text-align:center;padding:20px}.top h1{font-size:42px;margin:10px 0;background:linear-gradient(90deg,var(--primary),var(--accent),var(--secondary));-webkit-background-clip:text;background-clip:text;color:transparent}
.themebar{display:flex;justify-content:center;gap:8px;flex-wrap:wrap;margin:12px 0 22px}
.themebar button{border:1px solid var(--border);background:var(--surface);color:var(--text);padding:8px 13px;border-radius:999px;font-size:12px;cursor:pointer}
.themebar button:hover{border-color:var(--primary)}
.download{background:linear-gradient(90deg,var(--secondary),var(--primary));color:#04111c;padding:13px 22px;border-radius:11px;text-decoration:none;display:inline-block;margin:8px;font-weight:800}
.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:25px 0}
.stat{background:var(--surface);border:1px solid var(--border);padding:23px;border-radius:18px;text-align:center;box-shadow:0 10px 30px #00000030}
.stat b{display:block;font-size:31px;color:var(--primary)}.stat span{color:var(--muted)}
.section,.slide{background:linear-gradient(145deg,#ffffff0b,#ffffff04);border:1px solid var(--border);padding:26px;border-radius:18px;margin:18px 0;box-shadow:0 10px 35px #00000025}
.slide h2{margin-top:0;color:var(--primary)}
.score{font-size:25px;font-weight:800;color:var(--accent)}
.bar{height:11px;background:#ffffff12;border-radius:10px;overflow:hidden;margin:10px 0 18px}.fill{height:100%;background:linear-gradient(90deg,var(--secondary),var(--primary),var(--accent));border-radius:10px}
.issue{background:#fb718512;border:1px solid #fb718550;padding:16px;border-radius:12px;margin-top:15px}.issue strong{color:var(--danger)}
.suggest{background:#34d39912;border:1px solid #34d39950;padding:16px;border-radius:12px;margin-top:15px}.suggest strong{color:var(--success)}
p{line-height:1.75}.content{background:#00000018;border:1px solid var(--border);padding:16px;border-radius:12px}.muted{color:var(--muted)}
@media(max-width:800px){.summary{grid-template-columns:1fr 1fr}}
@media(max-width:500px){.summary{grid-template-columns:1fr}.top h1{font-size:32px}}
</style>
<script>
function setTheme(name){
 const themes={
 ocean:{bg:"#08111f",surface:"#0f1f33",surface2:"#132942",text:"#eef7ff",muted:"#9bb0c7",primary:"#22d3ee",secondary:"#3b82f6",accent:"#a78bfa",border:"#23415f",success:"#34d399",danger:"#fb7185"},
 sunset:{bg:"#1b0c12",surface:"#2b141d",surface2:"#3a1b25",text:"#fff5ed",muted:"#d6aaa0",primary:"#fb7185",secondary:"#f97316",accent:"#facc15",border:"#63303a",success:"#4ade80",danger:"#ef4444"},
 emerald:{bg:"#071713",surface:"#0d241e",surface2:"#123329",text:"#effff7",muted:"#9cc8b5",primary:"#34d399",secondary:"#14b8a6",accent:"#a3e635",border:"#1d5545",success:"#4ade80",danger:"#fb7185"},
 royal:{bg:"#100b25",surface:"#1b1439",surface2:"#251b4e",text:"#f7f3ff",muted:"#b9acd8",primary:"#c084fc",secondary:"#6366f1",accent:"#f472b6",border:"#493b78",success:"#34d399",danger:"#fb7185"}
 };
 const t=themes[name]||themes.ocean;
 Object.entries(t).forEach(([k,v])=>document.documentElement.style.setProperty("--"+k,v));
 localStorage.setItem("pptTheme",name);
}
window.addEventListener("DOMContentLoaded",()=>setTheme(localStorage.getItem("pptTheme")||"ocean"));
</script>

</head>
<body>
<div class="container">
<div class="top">
<div class="badge">SLIDESENSE AI REPORT</div>
<h1>Presentation Analysis Report</h1>
<p class="muted">File: {{ filename }}</p>
<a class="download" href="/download/{{ pdf_name }}">Download Final PDF Report</a>
<a class="download" href="/">Analyze Another PPT</a>
<div class="themebar"><button type="button" onclick="setTheme('ocean')">Ocean</button><button type="button" onclick="setTheme('sunset')">Sunset</button><button type="button" onclick="setTheme('emerald')">Emerald</button><button type="button" onclick="setTheme('royal')">Royal Purple</button></div>
</div>

<div class="summary">
<div class="stat"><b>{{ overall_score }}%</b><span>Overall Score</span></div>
<div class="stat"><b>{{ total_slides }}</b><span>Total Slides</span></div>
<div class="stat"><b>{{ total_issues }}</b><span>Issues Found</span></div>
<div class="stat"><b>{{ total_suggestions }}</b><span>Suggestions</span></div>
</div>

<div class="section">
<h2>Overall Report</h2>
<p>{{ overall_paragraph }}</p>
</div>

<h1>Individual Slide Analysis</h1>

{% for s in slides %}
<div class="slide">
<h2>Slide {{ s.number }} — {{ s.title }}</h2>
<div class="score">Slide Score: {{ s.score }}%</div>
<div class="bar"><div class="fill" style="width:{{ s.score }}%"></div></div>

<div class="content">
<strong>Extracted Content</strong>
<p>{{ s.text if s.text else "No readable text was detected on this slide." }}</p>
</div>

<div class="issue">
<strong>Issues and Explanation</strong>
{% for issue in s.issues %}
<p>• {{ issue }}</p>
{% endfor %}
</div>

<div class="suggest">
<strong>Suggestions in Paragraph Format</strong>
<p>{{ s.suggestion_paragraph }}</p>
</div>
</div>
{% endfor %}

<div class="section">
<h2>Final Recommendation</h2>
<p>{{ final_paragraph }}</p>
</div>
</div>
</body>
</html>
"""

def clean_text(text):
    return re.sub(r"\s+", " ", text or "").strip()

def extract_slides(path):
    prs = Presentation(path)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        title = ""
        if slide.shapes.title and hasattr(slide.shapes.title, "text"):
            title = clean_text(slide.shapes.title.text)
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                t = clean_text(shape.text)
                if t and t != title:
                    texts.append(t)
        text = " ".join(texts)
        slides.append({
            "number": i,
            "title": title if title else f"Slide {i}",
            "text": text,
            "text_blocks": texts
        })
    return slides

def analyze_slide(slide):
    title = slide["title"].lower()
    text = slide["text"]
    lower = text.lower()
    issues = []
    score = 100

    if not text:
        issues.append("No readable textual content was detected. The slide may depend mainly on images, diagrams, or other visual elements.")
        score -= 35
    elif len(text) < 40:
        issues.append("The slide contains very limited textual information, so the topic may not be sufficiently explained.")
        score -= 20
    elif len(text) > 900:
        issues.append("The slide contains a large amount of text, which can reduce readability during a presentation.")
        score -= 10

    if not slide["title"] or slide["title"].startswith("Slide "):
        issues.append("A clear descriptive title is missing.")
        score -= 15

    if len(slide["text_blocks"]) <= 1 and text:
        issues.append("The content is presented as a single block instead of being divided into clear points or sections.")
        score -= 5

    if any(word in title for word in ["introduction", "about"]):
        if len(text) < 100:
            issues.append("The introduction does not provide enough background, motivation, or project context.")
            score -= 10

    if "objective" in title:
        if not any(w in lower for w in ["aim", "goal", "develop", "design", "improve", "provide"]):
            issues.append("The objectives are not clearly expressed as project goals.")
            score -= 15

    if any(word in title for word in ["methodology", "working", "workflow", "process"]):
        if len(slide["text_blocks"]) < 2:
            issues.append("The methodology or working process needs more step-by-step explanation.")
            score -= 12

    if any(word in title for word in ["result", "output", "performance"]):
        if not any(w in lower for w in ["result", "output", "accuracy", "performance", "observed", "achieved"]):
            issues.append("The slide does not clearly explain the obtained results or system performance.")
            score -= 15

    if "conclusion" in title and len(text) < 80:
        issues.append("The conclusion is too brief to summarize the project findings and outcomes.")
        score -= 12

    if "future" in title and len(text) < 60:
        issues.append("The future scope contains insufficient detail about possible improvements or extensions.")
        score -= 10

    if not issues:
        issues.append("No major content issue was detected by the automated checks. The slide should still be reviewed for technical correctness and visual design.")

    score = max(0, min(100, score))

    # Generate a paragraph instead of bullet-style suggestions.
    if not text:
        suggestion = (
            "This slide should be improved by adding a clear title and concise supporting content. "
            "If the slide is intended to communicate information through an image, chart, or diagram, "
            "add a short explanation describing what the visual represents and why it is important to the project."
        )
    elif len(text) > 900:
        suggestion = (
            "The content should be simplified and reorganized into shorter points or sections. "
            "Important information can be retained while repetitive explanations are removed. "
            "This will make the slide easier to read and allow the presenter to explain the details verbally."
        )
    elif score < 70:
        suggestion = (
            "This slide requires improvement before the final presentation. The title and content should be made more specific, "
            "and missing explanations should be added so that the audience can understand the purpose of the slide without confusion. "
            "The information should be organized logically, with important technical details presented clearly and unnecessary text avoided."
        )
    elif score < 90:
        suggestion = (
            "The slide is reasonably structured, but it can be improved by making the explanation more specific and easier to understand. "
            "Use a descriptive title, concise points, and relevant supporting information. Any technical terms, results, or process steps "
            "should be explained clearly so the slide contributes directly to the overall project story."
        )
    else:
        suggestion = (
            "The slide has a good basic content structure. To make it more effective, maintain a clear title, concise wording, "
            "consistent terminology, and relevant supporting information. The final presentation should also be checked for visual consistency, "
            "technical accuracy, readable fonts, and alignment with the project objectives."
        )

    return score, issues, suggestion

def build_overall_paragraph(slides, overall):
    weak = sum(1 for s in slides if s["score"] < 70)
    if overall >= 90:
        level = "strong"
    elif overall >= 75:
        level = "good"
    elif overall >= 60:
        level = "moderate"
    else:
        level = "needs substantial improvement"

    return (
        f"The uploaded presentation contains {len(slides)} slides and received an overall automated score of {overall}%. "
        f"Based on the available slide text and structural checks, the presentation currently shows a {level} level of content quality. "
        f"The analysis identified {sum(len(s['issues']) for s in slides)} issue observations across the slides, with {weak} slides scoring below 70%. "
        f"The main improvement areas are clarity of slide titles, completeness of explanations, logical organization, readability, and sufficient explanation "
        f"of objectives, methodology, results, conclusion, and future scope where those sections are present. "
        f"The slide-specific suggestions should be applied before producing the final academic or professional presentation."
    )

def build_final_paragraph(slides):
    scores = [s["score"] for s in slides]
    avg = sum(scores) / len(scores) if scores else 0
    return (
        f"The final review indicates an average slide score of {avg:.1f}%. "
        "The presentation should be revised by addressing the individual issues identified for each slide, "
        "strengthening explanations where information is incomplete, and keeping the content concise and logically connected. "
        "After these improvements, the presentation should be checked manually for technical accuracy, grammar, visual consistency, "
        "font readability, image quality, references, and alignment with the actual project implementation. "
        "This automated report is intended as a structured review and should be used together with human review before submission."
    )

def pdf_escape(text):
    return html.escape(str(text)).replace("\n", "<br/>")

def generate_pdf(filename, slides, overall, overall_paragraph, final_paragraph):
    pdf_path = os.path.join(REPORT_FOLDER, filename)
    doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        rightMargin=18*mm, leftMargin=18*mm,
        topMargin=18*mm, bottomMargin=18*mm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleX", parent=styles["Title"], fontSize=21, alignment=TA_CENTER, spaceAfter=18)
    h_style = ParagraphStyle("HX", parent=styles["Heading2"], fontSize=14, spaceBefore=12, spaceAfter=7)
    body = ParagraphStyle("BodyX", parent=styles["BodyText"], fontSize=9.8, leading=14, spaceAfter=8)
    small = ParagraphStyle("SmallX", parent=body, fontSize=8.5)

    story = [
        Paragraph("PPT Analysis and Automated Report", title_style),
        Paragraph("Presentation Quality Review", styles["Heading2"]),
        Spacer(1, 10),
        Paragraph(pdf_escape(overall_paragraph), body),
        Spacer(1, 10)
    ]

    data = [
        [Paragraph("<b>Metric</b>", small), Paragraph("<b>Value</b>", small)],
        [Paragraph("Overall Score", small), Paragraph(f"{overall}%", small)],
        [Paragraph("Total Slides", small), Paragraph(str(len(slides)), small)],
        [Paragraph("Total Issue Observations", small), Paragraph(str(sum(len(s["issues"]) for s in slides)), small)]
    ]
    table = Table(data, colWidths=[85*mm, 55*mm])
    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6)
    ]))
    story.append(table)
    story.append(PageBreak())

    story.append(Paragraph("Individual Slide Analysis", title_style))

    for s in slides:
        story.append(Paragraph(f"Slide {s['number']}: {pdf_escape(s['title'])}", h_style))
        story.append(Paragraph(f"<b>Slide Score:</b> {s['score']}%", body))
        story.append(Paragraph("<b>Extracted Content:</b> " + pdf_escape(s["text"] if s["text"] else "No readable text detected."), body))
        story.append(Paragraph("<b>Issues and Explanation:</b>", body))
        for issue in s["issues"]:
            story.append(Paragraph("• " + pdf_escape(issue), body))
        story.append(Paragraph("<b>Suggestions:</b>", body))
        story.append(Paragraph(pdf_escape(s["suggestion_paragraph"]), body))
        story.append(Spacer(1, 8))

    story.append(PageBreak())
    story.append(Paragraph("Final Recommendation", title_style))
    story.append(Paragraph(pdf_escape(final_paragraph), body))
    doc.build(story)
    return pdf_path

@app.route("/")
def index():
    return render_template_string(INDEX_HTML)

@app.route("/analyze", methods=["POST"])
def analyze():
    file = request.files.get("ppt_file")
    if not file or not file.filename:
        return render_template_string(INDEX_HTML) + '<div class="error">Please select a PPT or PPTX file.</div>'
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".ppt", ".pptx"]:
        return "Only .ppt and .pptx files are supported.", 400

    safe = re.sub(r"[^A-Za-z0-9_.-]", "_", file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, safe)
    file.save(input_path)

    try:
        slides = extract_slides(input_path)
    except Exception as e:
        return f"Could not read the PowerPoint file. Error: {html.escape(str(e))}", 400

    analyzed = []
    for slide in slides:
        score, issues, suggestion = analyze_slide(slide)
        analyzed.append({
            **slide,
            "score": score,
            "issues": issues,
            "suggestion_paragraph": suggestion
        })

    overall = round(sum(s["score"] for s in analyzed) / len(analyzed)) if analyzed else 0
    overall_paragraph = build_overall_paragraph(analyzed, overall)
    final_paragraph = build_final_paragraph(analyzed)

    base = os.path.splitext(safe)[0]
    pdf_name = base + "_analysis_report.pdf"
    generate_pdf(pdf_name, analyzed, overall, overall_paragraph, final_paragraph)

    return render_template_string(
        RESULT_HTML,
        filename=safe,
        pdf_name=pdf_name,
        overall_score=overall,
        total_slides=len(analyzed),
        total_issues=sum(len(s["issues"]) for s in analyzed),
        total_suggestions=sum(len(s["issues"]) for s in analyzed),
        overall_paragraph=overall_paragraph,
        final_paragraph=final_paragraph,
        slides=analyzed
    )

@app.route("/download/<filename>")
def download(filename):
    safe = os.path.basename(filename)
    path = os.path.join(REPORT_FOLDER, safe)
    if not os.path.exists(path):
        return "Report not found.", 404
    return send_file(path, as_attachment=True, download_name=safe)

if __name__ == "__main__":
    print("\nPPT Analyzer is running.")
    print("Open: http://127.0.0.1:5000\n")
    app.run(debug=True)
