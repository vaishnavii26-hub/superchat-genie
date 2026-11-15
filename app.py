import os
import io
import json
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv
from google import genai
from google.genai import types
import docx
from pypdf import PdfReader
from fpdf import FPDF

load_dotenv()

app = Flask(__name__)
DB_PATH = "documents.db"

# Initialize Gemini Client
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None
MODEL_ID = "gemini-3.5-flash-lite"


# ---------------- DATABASE SETUP ----------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            content TEXT DEFAULT '',
            folder TEXT DEFAULT 'General',
            updated_at TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_id INTEGER NOT NULL,
            content TEXT NOT NULL,
            saved_at TEXT NOT NULL,
            FOREIGN KEY (doc_id) REFERENCES documents (id) ON DELETE CASCADE
        )
    """)
    # Insert starter welcome doc if empty
    count = c.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
    if count == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        starter_html = """<h1>Welcome to AuraWrite AI Studio</h1>
<p>Start typing here, paste an article, or upload a <b>DOCX</b>, <b>PDF</b>, or <b>TXT</b> file from the top bar.</p>
<p>Try selecting any paragraph and using the right-hand panel to check grammar, adjust tone, analyze readability, or engineer better AI prompts.</p>"""
        c.execute(
            "INSERT INTO documents (title, content, folder, updated_at) VALUES (?, ?, ?, ?)",
            ("Getting Started Guide", starter_html, "General", now),
        )
    conn.commit()
    conn.close()


init_db()


# ---------------- AI HELPER ----------------
def ask_gemini(prompt, system_instruction="", json_mode=False):
    if not client:
        raise ValueError("GEMINI_API_KEY is missing. Add it to your .env file or server environment variables.")
    config_args = {
        "system_instruction": system_instruction or "You are an expert professional writing assistant and editor.",
        "temperature": 0.4 if json_mode else 0.7,
    }
    if json_mode:
        config_args["response_mime_type"] = "application/json"

    response = client.models.generate_content(
        model=MODEL_ID,
        contents=prompt,
        config=types.GenerateContentConfig(**config_args),
    )
    return response.text


# ---------------- FRONTEND ROUTE ----------------
@app.route("/")
def index():
    return render_template("index.html")


# ---------------- DOCUMENT CRUD & HISTORY ----------------
@app.route("/api/docs", methods=["GET"])
def list_docs():
    conn = get_db()
    docs = conn.execute("SELECT * FROM documents ORDER BY updated_at DESC").fetchall()
    conn.close()
    return jsonify([dict(d) for d in docs])


@app.route("/api/docs", methods=["POST"])
def create_doc():
    data = request.json or {}
    title = data.get("title", "Untitled Document")
    content = data.get("content", "<p></p>")
    folder = data.get("folder", "General")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "INSERT INTO documents (title, content, folder, updated_at) VALUES (?, ?, ?, ?)",
        (title, content, folder, now),
    )
    doc_id = c.lastrowid
    c.execute(
        "INSERT INTO versions (doc_id, content, saved_at) VALUES (?, ?, ?)",
        (doc_id, content, now),
    )
    conn.commit()
    doc = conn.execute("SELECT * FROM documents WHERE id = ?", (doc_id,)).fetchone()
    conn.close()
    return jsonify(dict(doc))


@app.route("/api/docs/<int:doc_id>", methods=["PUT"])
def update_doc(doc_id):
    data = request.json or {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    c = conn.cursor()
    existing = c.execute("SELECT * FROM documents WHERE id = ?", (doc_id,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({"error": "Document not found"}), 404

    title = data.get("title", existing["title"])
    content = data.get("content", existing["content"])
    folder = data.get("folder", existing["folder"])
    save_snapshot = data.get("save_snapshot", False)

    c.execute(
        "UPDATE documents SET title = ?, content = ?, folder = ?, updated_at = ? WHERE id = ?",
        (title, content, folder, now, doc_id),
    )
    if save_snapshot:
        c.execute(
            "INSERT INTO versions (doc_id, content, saved_at) VALUES (?, ?, ?)",
            (doc_id, content, now),
        )
    conn.commit()
    conn.close()
    return jsonify({"status": "saved", "updated_at": now})


@app.route("/api/docs/<int:doc_id>", methods=["DELETE"])
def delete_doc(doc_id):
    conn = get_db()
    conn.execute("DELETE FROM versions WHERE doc_id = ?", (doc_id,))
    conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
    conn.commit()
    conn.close()
    return jsonify({"status": "deleted"})


@app.route("/api/docs/<int:doc_id>/duplicate", methods=["POST"])
def duplicate_doc(doc_id):
    conn = get_db()
    c = conn.cursor()
    doc = c.execute("SELECT * FROM documents WHERE id = ?", (doc_id,)).fetchone()
    if not doc:
        conn.close()
        return jsonify({"error": "Not found"}), 404
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute(
        "INSERT INTO documents (title, content, folder, updated_at) VALUES (?, ?, ?, ?)",
        (f"{doc['title']} (Copy)", doc["content"], doc["folder"], now),
    )
    new_id = c.lastrowid
    conn.commit()
    new_doc = c.execute("SELECT * FROM documents WHERE id = ?", (new_id,)).fetchone()
    conn.close()
    return jsonify(dict(new_doc))


@app.route("/api/docs/<int:doc_id>/versions", methods=["GET"])
def get_versions(doc_id):
    conn = get_db()
    vers = conn.execute(
        "SELECT * FROM versions WHERE doc_id = ? ORDER BY id DESC LIMIT 15", (doc_id,)
    ).fetchall()
    conn.close()
    return jsonify([dict(v) for v in vers])


# ---------------- FILE UPLOAD & EXPORT ----------------
@app.route("/api/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    filename = file.filename.lower()
    extracted_html = ""

    try:
        if filename.endswith(".txt"):
            text = file.read().decode("utf-8", errors="ignore")
            paragraphs = [f"<p>{line.strip()}</p>" for line in text.splitlines() if line.strip()]
            extracted_html = "".join(paragraphs) or "<p></p>"

        elif filename.endswith(".docx"):
            doc = docx.Document(file)
            paragraphs = [f"<p>{p.text}</p>" for p in doc.paragraphs if p.text.strip()]
            extracted_html = "".join(paragraphs) or "<p></p>"

        elif filename.endswith(".pdf"):
            reader = PdfReader(file)
            pages_text = []
            for page in reader.pages:
                txt = page.extract_text()
                if txt:
                    for line in txt.splitlines():
                        if line.strip():
                            pages_text.append(f"<p>{line.strip()}</p>")
            extracted_html = "".join(pages_text) or "<p>No extractable text found in PDF.</p>"
        else:
            return jsonify({"error": "Unsupported file format. Please upload DOCX, PDF, or TXT."}), 400

        title = os.path.splitext(file.filename)[0]
        return jsonify({"title": title, "content": extracted_html})
    except Exception as e:
        return jsonify({"error": f"Failed to parse file: {str(e)}"}), 500


@app.route("/api/export", methods=["POST"])
def export_file():
    data = request.json or {}
    fmt = data.get("format", "txt")
    title = data.get("title", "document")
    plain_text = data.get("text", "")

    if fmt == "txt":
        buf = io.BytesIO(plain_text.encode("utf-8"))
        return send_file(buf, as_attachment=True, download_name=f"{title}.txt", mimetype="text/plain")

    elif fmt == "docx":
        doc = docx.Document()
        doc.add_heading(title, 0)
        for line in plain_text.splitlines():
            if line.strip():
                doc.add_paragraph(line.strip())
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{title}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    elif fmt == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Helvetica", "B", 16)
        safe_title = title.encode("latin-1", "replace").decode("latin-1")
        pdf.cell(0, 10, safe_title, ln=True)
        pdf.ln(4)
        pdf.set_font("Helvetica", size=11)
        safe_text = plain_text.encode("latin-1", "replace").decode("latin-1")
        pdf.multi_cell(0, 7, safe_text)
        buf = io.BytesIO(pdf.output())
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=f"{title}.pdf", mimetype="application/pdf")

    return jsonify({"error": "Invalid format"}), 400


# ---------------- AI FEATURES ----------------
@app.route("/api/ai/grammar", methods=["POST"])
def ai_grammar():
    text = (request.json or {}).get("text", "").strip()
    if not text:
        return jsonify({"error": "Please enter some text first."}), 400

    prompt = f"""Analyze the following text for grammar, spelling, punctuation, subject-verb agreement, incorrect tenses, repeated words, missing words, improper capitalization, sentence fragments, and run-on sentences.
Return a JSON object with a single key "corrections" containing a list of objects. Each object must have:
- "original": the exact problematic substring from the text
- "corrected": the corrected replacement
- "category": e.g., "Grammar", "Spelling", "Punctuation", "Sentence Structure"
- "explanation": concise explanation of why the correction is needed

If there are no errors, return {{"corrections": []}}.

Text:
{text}"""
    try:
        raw = ask_gemini(prompt, "You are a strict, precise copyeditor.", json_mode=True)
        return jsonify(json.loads(raw))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/enhance", methods=["POST"])
def ai_enhance():
    data = request.json or {}
    text = data.get("text", "").strip()
    action = data.get("action", "Improve clarity")
    tone = data.get("tone", "Neutral")

    if not text:
        return jsonify({"error": "Please provide text in the editor or select a passage."}), 400

    prompt = f"""Task: {action}
Target Tone: {tone}

Instructions:
- Preserve the original meaning, facts, and context.
- Return clean HTML using <p>, <ul>, <li>, <b>, <i> tags where appropriate (do NOT wrap in markdown ```html blocks).
- Do not add conversational filler before or after the rewritten content.

Original Text:
{text}"""
    try:
        result = ask_gemini(prompt, "You are an expert developmental editor and style specialist.")
        cleaned = result.replace("```html", "").replace("```", "").strip()
        return jsonify({"result": cleaned})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/suggestions", methods=["POST"])
def ai_suggestions():
    text = (request.json or {}).get("text", "").strip()
    if not text:
        return jsonify({"error": "Write some content first to receive editorial suggestions."}), 400

    prompt = f"""Act as an intelligent writing advisor. Analyze the text and provide 4 to 6 high-impact editorial suggestions covering stronger vocabulary, missing information, hook/conclusion improvements, stronger arguments, and flow.
Return a JSON object with key "suggestions" containing a list of objects with:
- "title": short title of the suggestion
- "type": e.g., "Vocabulary", "Structure", "Missing Info", "Hook & Intro", "Argument"
- "explanation": why this improves the piece
- "suggested_addition": ready-to-insert HTML snippet (<p>...</p>) that implements the improvement

Text:
{text}"""
    try:
        raw = ask_gemini(prompt, "You are a senior publishing editor.", json_mode=True)
        return jsonify(json.loads(raw))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/generate", methods=["POST"])
def ai_generate():
    data = request.json or {}
    doc_type = data.get("doc_type", "Article")
    topic = data.get("topic", "")
    audience = data.get("audience", "General readers")
    tone = data.get("tone", "Professional")
    length = data.get("length", "Medium (300-500 words)")
    points = data.get("points", "")

    if not topic:
        return jsonify({"error": "Please enter a topic or instructions."}), 400

    prompt = f"""Write a complete, polished {doc_type}.
Topic/Purpose: {topic}
Target Audience: {audience}
Tone: {tone}
Target Length: {length}
Key Points to Include: {points}

Format the output in clean semantic HTML (<h2>, <h3>, <p>, <ul>, <li>, <b>). Do NOT include markdown code fences."""
    try:
        result = ask_gemini(prompt, "You are an elite professional content writer.")
        cleaned = result.replace("```html", "").replace("```", "").strip()
        return jsonify({"result": cleaned})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/improve-prompt", methods=["POST"])
def ai_improve_prompt():
    raw_prompt = (request.json or {}).get("prompt", "").strip()
    if not raw_prompt:
        return jsonify({"error": "Please enter a basic prompt to improve."}), 400

    prompt = f"""Analyze this rough user prompt and transform it into a high-precision, production-grade AI prompt.
Return a JSON object with:
- "missing_elements": list of strings identifying missing instructions, context, or constraints
- "role": recommended AI persona/role
- "improved_prompt": the complete, structured, ready-to-use prompt (with Role, Context, Task, Constraints, Output Format, and Example)
- "explanation": brief breakdown of why the new prompt works better
- "variations": list of 2 alternative prompt variations (one concise, one ultra-detailed)

User's Basic Prompt:
{raw_prompt}"""
    try:
        raw = ask_gemini(prompt, "You are a world-class prompt engineer.", json_mode=True)
        return jsonify(json.loads(raw))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/analyze", methods=["POST"])
def ai_analyze():
    text = (request.json or {}).get("text", "").strip()
    if not text:
        return jsonify({"error": "Add text to analyze readability and writing quality."}), 400

    prompt = f"""Perform a comprehensive linguistic and readability analysis of the text below.
Return a JSON object with:
- "readability_score": integer 0-100 (100 = easiest to read)
- "reading_level": e.g., "Grade 8 (Accessible)", "University / Technical"
- "clarity_score": integer 0-100
- "grammar_score": integer 0-100
- "vocabulary_complexity": "Basic", "Moderate", or "Advanced"
- "tone_detected": e.g., "Professional & Informative"
- "active_voice_pct": integer 0-100
- "passive_voice_pct": integer 0-100
- "repeated_words": list of up to 5 overused words
- "actionable_tips": list of 4 specific tips to improve readability, sentence variety, and clarity

Text:
{text}"""
    try:
        raw = ask_gemini(prompt, "You are a computational linguist and readability analyst.", json_mode=True)
        return jsonify(json.loads(raw))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/ai/tool", methods=["POST"])
def ai_tool():
    data = request.json or {}
    text = data.get("text", "").strip()
    tool = data.get("tool", "summarize")
    extra = data.get("extra", "")

    if not text:
        return jsonify({"error": "Please enter text in the editor first."}), 400

    instructions = {
        "summarize": "Provide an executive summary with a concise overview paragraph followed by key bullet points.",
        "outline": "Generate a structured, hierarchical content outline (Roman numerals / headings and sub-points) based on this text.",
        "titles": "Suggest 8 compelling, click-worthy titles and section headings suited for this content.",
        "translate": f"Translate the text accurately into {extra or 'Spanish'}, preserving tone and formatting.",
        "citations": f"Format references and extract/generate properly formatted citations in {extra or 'APA 7th Edition'} style.",
        "originality": "Analyze the text for generic AI-sounding clichés, overly common phrasing, and originality risks, and provide a uniqueness assessment with tips to make the voice more distinctive.",
        "examples": "Generate 3 concrete real-world examples or case studies that illustrate the main points in this text.",
    }

    task_desc = instructions.get(tool, "Improve and analyze this text.")
    prompt = f"{task_desc}\n\nReturn clean HTML (<p>, <ul>, <li>, <h3>, <b>) without markdown fences.\n\nText:\n{text}"

    try:
        result = ask_gemini(prompt, "You are an advanced AI writing assistant.")
        cleaned = result.replace("```html", "").replace("```", "").strip()
        return jsonify({"result": cleaned})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)