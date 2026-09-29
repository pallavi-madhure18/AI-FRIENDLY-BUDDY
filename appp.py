import os
import re
import html
import hashlib
import sqlite3
from pathlib import Path
from datetime import datetime
from collections import Counter

import streamlit as st

# Gemini SDK
try:
    from google import genai
    from google.genai import types
except Exception:
    genai = None
    types = None

# Optional document readers
try:
    import PyPDF2
except Exception:
    PyPDF2 = None

try:
    from docx import Document
except Exception:
    Document = None

try:
    import pandas as pd
except Exception:
    pd = None


# ============================================================
# PAGE
# ============================================================
st.set_page_config(
    page_title="StudyBuddy AI",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "study_data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_FILE = DATA_DIR / "studybuddy.db"
CSS_FILE = BASE_DIR / "style.css"


# ============================================================
# CSS
# ============================================================
if CSS_FILE.exists():
    css = CSS_FILE.read_text(encoding="utf-8")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


# ============================================================
# DATABASE
# ============================================================
def get_db():
    conn = sqlite3.connect(str(DB_FILE), timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question TEXT NOT NULL,
                answer TEXT NOT NULL,
                mode TEXT,
                file_name TEXT,
                created_at TEXT NOT NULL,
                favorite INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                file_hash TEXT UNIQUE NOT NULL,
                content TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            )
        """)


init_db()


def save_chat(question, answer, mode, file_name):
    with get_db() as conn:
        conn.execute(
            """
            INSERT INTO history
            (question, answer, mode, file_name, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                question,
                answer,
                mode,
                file_name or "",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )


def get_history(limit=None):
    with get_db() as conn:
        if limit is not None:
            return conn.execute(
                "SELECT * FROM history ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return conn.execute(
            "SELECT * FROM history ORDER BY id DESC"
        ).fetchall()


def delete_chat(chat_id):
    with get_db() as conn:
        conn.execute("DELETE FROM history WHERE id = ?", (chat_id,))


def clear_history():
    with get_db() as conn:
        conn.execute("DELETE FROM history")


def toggle_favorite(chat_id):
    with get_db() as conn:
        conn.execute(
            """
            UPDATE history
            SET favorite = CASE WHEN favorite = 1 THEN 0 ELSE 1 END
            WHERE id = ?
            """,
            (chat_id,),
        )


def get_favorites():
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM history WHERE favorite = 1 ORDER BY id DESC"
        ).fetchall()


def save_note(title, content):
    with get_db() as conn:
        conn.execute(
            """
            INSERT INTO notes (title, content, created_at)
            VALUES (?, ?, ?)
            """,
            (
                title,
                content,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )


def get_notes():
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM notes ORDER BY id DESC"
        ).fetchall()


def delete_note(note_id):
    with get_db() as conn:
        conn.execute("DELETE FROM notes WHERE id = ?", (note_id,))


def save_document(file_name, file_hash, content):
    with get_db() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO documents
            (file_name, file_hash, content, uploaded_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                file_name,
                file_hash,
                content,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )


def get_document(file_hash):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM documents WHERE file_hash = ?",
            (file_hash,),
        ).fetchone()


# ============================================================
# SESSION STATE
# ============================================================
DEFAULTS = {
    "document_text": "",
    "document_name": "",
    "document_hash": "",
    "last_answer": "",
    "question_box": "",
    "selected_history_answer": "",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# FILE READER
# ============================================================
def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith((".txt", ".md")):
        return uploaded_file.getvalue().decode("utf-8", errors="ignore"), None

    if name.endswith(".pdf"):
        if PyPDF2 is None:
            return None, "PDF support is unavailable. Install PyPDF2."
        try:
            reader = PyPDF2.PdfReader(uploaded_file)
            pages = [(page.extract_text() or "") for page in reader.pages]
            text = "\n".join(pages).strip()
            if not text:
                return None, (
                    "This PDF has no selectable text. It may be a scanned/image-only PDF."
                )
            return text, None
        except Exception as exc:
            return None, f"Could not read the PDF: {exc}"

    if name.endswith(".docx"):
        if Document is None:
            return None, "DOCX support is unavailable. Install python-docx."
        try:
            doc = Document(uploaded_file)
            parts = [p.text for p in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts).strip(), None
        except Exception as exc:
            return None, f"Could not read the Word file: {exc}"

    if name.endswith(".csv"):
        if pd is None:
            return None, "CSV support is unavailable. Install pandas."
        try:
            df = pd.read_csv(uploaded_file)
            return df.to_string(index=False), None
        except Exception as exc:
            return None, f"Could not read CSV: {exc}"

    if name.endswith((".xlsx", ".xls")):
        if pd is None:
            return None, "Excel support is unavailable. Install pandas and openpyxl."
        try:
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            output = []
            for sheet_name, df in sheets.items():
                output.append(f"Sheet: {sheet_name}\n{df.to_string(index=False)}")
            return "\n\n".join(output), None
        except Exception as exc:
            return None, f"Could not read Excel file: {exc}"

    return None, "This file type is not supported."


# ============================================================
# GEMINI CONFIGURATION
# ============================================================
DEFAULT_MODEL = "gemini-3.8-flash"


def get_secret_or_env(name):
    """Read a secret from Streamlit Cloud/local secrets, then environment."""
    try:
        value = st.secrets.get(name)
        if value:
            return str(value).strip()
    except Exception:
        pass

    value = os.getenv(name, "")
    return value.strip()


def get_api_key():
    # GEMINI_API_KEY is the primary name.
    # GOOGLE_API_KEY is also accepted for convenience.
    return get_secret_or_env("GEMINI_API_KEY") or get_secret_or_env("GOOGLE_API_KEY")


def get_model_name():
    model = get_secret_or_env("GEMINI_MODEL")
    return model or DEFAULT_MODEL


def get_gemini_client():
    if genai is None:
        raise RuntimeError(
            "The Gemini package is not installed. Run: pip install -r requirements.txt"
        )

    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "No Gemini API key was found. Add GEMINI_API_KEY to "
            ".streamlit/secrets.toml locally or Streamlit Cloud → Settings → Secrets."
        )

    try:
        return genai.Client(api_key=api_key)
    except Exception as exc:
        raise RuntimeError(f"Could not initialize Gemini: {exc}") from exc


def friendly_ai_error(exc):
    message = str(exc)
    lower = message.lower()

    if "api key" in lower or "api_key" in lower or "unauthenticated" in lower:
        return (
            "Gemini rejected the API key. Create a fresh key in Google AI Studio "
            "and replace GEMINI_API_KEY in your local secrets or Streamlit Cloud Secrets."
        )

    if "quota" in lower or "resource_exhausted" in lower or "429" in lower:
        return (
            "Gemini rate/quota limit was reached. Wait and try again, or check "
            "your Google AI Studio/Cloud quota."
        )

    if "not found" in lower or "model" in lower and "404" in lower:
        return (
            f"The configured Gemini model '{get_model_name()}' is unavailable. "
            "Use a currently supported model such as gemini-3.8-flash."
        )

    if "permission" in lower or "forbidden" in lower or "403" in lower:
        return (
            "Google rejected access to the Gemini API. Check the API key, "
            "project access, and API availability in Google AI Studio."
        )

    return f"Gemini request failed: {message}"


def ask_ai(question, mode, document_text=""):
    mode_instructions = {
        "💬 Friendly Tutor": "Teach clearly, patiently and simply. Use examples.",
        "📝 Exam Prep": "Focus on exam-ready facts, important questions, definitions and concise answers.",
        "🧩 Quiz Master": "Create useful quiz questions and answers based on the topic or uploaded material.",
        "🧠 Explain Simply": "Explain the topic for a beginner using easy language and analogies.",
        "⚡ Quick Revision": "Create a compact revision sheet with key facts, formulas and memory tricks.",
        "🎯 Study Planner": "Create a realistic study plan with priorities, sessions, breaks and revision.",
        "💡 Doubt Solver": "Solve the student's doubt step by step and explain common mistakes.",
    }

    # Keep a reasonable context size. The current Gemini model supports a much
    # larger context, but limiting pasted document text keeps requests practical.
    context = document_text[:50000] if document_text else "No study file is currently uploaded."

    system_prompt = f"""
You are StudyBuddy AI, a friendly study assistant.

STUDY MODE:
{mode}

MODE INSTRUCTIONS:
{mode_instructions.get(mode, mode_instructions["💬 Friendly Tutor"])}

RULES:
1. Answer the student's actual question directly.
2. Use simple, readable language.
3. Use headings, bullets, examples and steps when useful.
4. For maths/science, show important calculation steps.
5. If the question is about the uploaded study material, use that material as the primary source.
6. Never invent facts from an uploaded file.
7. If the uploaded material does not contain the requested information, say so clearly.
8. Encourage understanding, not unexplained copying.
9. Keep the response organized and student-friendly.

UPLOADED STUDY MATERIAL:
{context}
""".strip()

    client = get_gemini_client()
    model = get_model_name()

    try:
        config = types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.4,
            max_output_tokens=4096,
        )
        response = client.models.generate_content(
            model=model,
            contents=question,
            config=config,
        )
    except Exception as exc:
        raise RuntimeError(friendly_ai_error(exc)) from exc

    answer = getattr(response, "text", None)
    if not answer:
        raise RuntimeError(
            "Gemini returned an empty response. Please try the question again."
        )
    return answer.strip()


def ai_configured():
    return bool(get_api_key()) and genai is not None


# ============================================================
# HELPERS
# ============================================================
def safe_html(text):
    return html.escape(str(text)).replace("\n", "<br>")


def render_answer(answer):
    st.markdown('<div class="answer-top">🤖 StudyBuddy AI</div>', unsafe_allow_html=True)
    st.markdown(answer)
    st.markdown(
        '<div class="answer-bottom">✨ Keep learning. You’ve got this!</div>',
        unsafe_allow_html=True,
    )


def run_ai_request(prompt, mode):
    try:
        with st.spinner("🤖 StudyBuddy is thinking..."):
            answer = ask_ai(
                prompt,
                mode,
                st.session_state.document_text,
            )
        save_chat(
            prompt,
            answer,
            mode,
            st.session_state.document_name,
        )
        st.session_state.last_answer = answer
        return answer
    except Exception as exc:
        st.error(f"⚠️ {exc}")
        return None


# ============================================================
# HERO
# ============================================================
st.markdown(
    """
    <div class="hero">
        <div class="robot">🤖</div>
        <div class="eyebrow">✦ AI POWERED PERSONAL STUDY SPACE ✦</div>
        <h1>Study<span>Buddy</span> AI</h1>
        <p>Ask questions • Upload notes • Revise • Practice • Learn 🚀</p>
        <div class="emoji-row">🧠 ✨ 📚 🎯 💡 📝 🚀</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🧭 Study Hub")

    mode = st.selectbox(
        "🎓 Study mode",
        [
            "💬 Friendly Tutor",
            "📝 Exam Prep",
            "🧩 Quiz Master",
            "🧠 Explain Simply",
            "⚡ Quick Revision",
            "🎯 Study Planner",
            "💡 Doubt Solver",
        ],
    )

    st.markdown("---")
    st.markdown("### 🔎 Search History")

    history_search = st.text_input(
        "Search your previous questions",
        placeholder="e.g. Python, physics, DBMS...",
        label_visibility="collapsed",
    )

    if history_search.strip():
        with get_db() as conn:
            term = f"%{history_search.strip()}%"
            searched = conn.execute(
                """
                SELECT * FROM history
                WHERE question LIKE ? OR answer LIKE ?
                ORDER BY id DESC
                """,
                (term, term),
            ).fetchall()

        if searched:
            st.caption(f"Found {len(searched)} result(s)")
            for row in searched[:10]:
                if st.button(
                    f"💬 {row['question'][:45]}",
                    key=f"search_{row['id']}",
                    use_container_width=True,
                ):
                    st.session_state.question_box = row["question"]
                    st.session_state.selected_history_answer = row["answer"]
                    st.rerun()
        else:
            st.info("No matching questions found.")

    st.markdown("---")
    st.markdown("### 📂 Study Material")

    uploaded = st.file_uploader(
        "Upload your notes",
        type=["pdf", "docx", "txt", "md", "csv", "xlsx", "xls"],
        help="Upload notes and ask StudyBuddy questions about them.",
    )

    if uploaded is not None:
        file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()

        if st.session_state.document_hash != file_hash:
            cached = get_document(file_hash)

            if cached:
                st.session_state.document_text = cached["content"]
                st.session_state.document_name = cached["file_name"]
                st.session_state.document_hash = file_hash
                st.success("⚡ Loaded from local storage.")
            else:
                with st.spinner("📖 Reading your file..."):
                    text, error = read_uploaded_file(uploaded)

                if error:
                    st.error(error)
                else:
                    st.session_state.document_text = text or ""
                    st.session_state.document_name = uploaded.name
                    st.session_state.document_hash = file_hash
                    save_document(
                        uploaded.name,
                        file_hash,
                        st.session_state.document_text,
                    )
                    st.success("✅ File read successfully!")

    if st.session_state.document_text:
        st.markdown(
            f"""
            <div class="file-card">
                📄 <b>{safe_html(st.session_state.document_name)}</b><br>
                <span>Ready for questions ✨</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🗑️ Remove current file", use_container_width=True):
            st.session_state.document_text = ""
            st.session_state.document_name = ""
            st.session_state.document_hash = ""
            st.rerun()

    st.markdown("---")
    st.markdown("### 🕘 History")

    all_history = get_history()
    favorites = get_favorites()

    st.metric("Questions asked", len(all_history))

    if st.button("🧹 Clear ALL History", use_container_width=True):
        clear_history()
        st.success("History cleared!")
        st.rerun()

    st.caption(f"⭐ Favorites: {len(favorites)}")


# ============================================================
# AI STATUS
# ============================================================
if genai is None:
    st.markdown(
        '<div class="status offline">🔴 Gemini package is missing • Run pip install -r requirements.txt</div>',
        unsafe_allow_html=True,
    )
elif ai_configured():
    st.markdown(
        f'<div class="status online">🟢 AI READY • {html.escape(get_model_name())}</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="status offline">🔴 AI KEY NOT CONFIGURED • Add GEMINI_API_KEY to your secrets</div>',
        unsafe_allow_html=True,
    )
    with st.expander("🔑 One-time API key setup"):
        st.markdown(
            """
**Local computer**

Create `.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "YOUR_REAL_GEMINI_API_KEY"
```

**Streamlit Cloud**

Open your app → **Settings / Secrets** and add the same line.

Never put the real API key in `app.py` or GitHub.
""",
        )


# ============================================================
# DASHBOARD
# ============================================================
history = get_history()
notes = get_notes()
favorites = get_favorites()

c1, c2, c3, c4 = st.columns(4)
c1.metric("💬 Asked", len(history))
c2.metric("⭐ Favorites", len(favorites))
c3.metric("📝 Saved Notes", len(notes))
c4.metric("📚 Study File", "Ready" if st.session_state.document_text else "None")


# ============================================================
# MOST ASKED
# ============================================================
st.markdown("### 🔥 Mostly Asked Questions")

if history:
    normalized = []
    for row in history:
        q = re.sub(r"\s+", " ", row["question"].strip().lower())
        normalized.append(q)

    popular = Counter(normalized).most_common(5)
    popular_cols = st.columns(len(popular))

    for index, (question_text, count) in enumerate(popular):
        with popular_cols[index]:
            if st.button(
                f"🔥 {question_text[:42]}",
                key=f"popular_{index}",
                use_container_width=True,
            ):
                st.session_state.question_box = question_text
                st.rerun()
            st.caption(f"Asked {count} time{'s' if count != 1 else ''}")
else:
    st.info("Your most-asked questions will appear here after you start studying.")


# ============================================================
# QUICK TOOLS
# ============================================================
st.markdown("### ⚡ Quick Study Tools")

q1, q2, q3, q4 = st.columns(4)

quick_actions = [
    (q1, "📌 Summarize", "Summarize my uploaded study material into clear exam-ready notes."),
    (q2, "❓ Make Quiz", "Create 10 useful quiz questions from my uploaded study material and provide the answers."),
    (q3, "🧠 Flashcards", "Turn my uploaded study material into concise question-and-answer flashcards."),
    (q4, "🔑 Key Points", "Extract the most important definitions, formulas, facts and concepts from my uploaded study material."),
]

for col, label, prompt in quick_actions:
    with col:
        if st.button(label, use_container_width=True):
            if not st.session_state.document_text:
                st.warning("📂 Upload study material first.")
            elif not ai_configured():
                st.error("🔴 Add a valid GEMINI_API_KEY first.")
            else:
                answer = run_ai_request(prompt, mode)
                if answer:
                    st.markdown("### 🌟 StudyBuddy Result")
                    render_answer(answer)


# ============================================================
# MAIN QUESTION
# ============================================================
st.markdown("### 💭 Ask StudyBuddy")

question = st.text_area(
    "Your question",
    placeholder=(
        "Type your question here...\n\n"
        "Example: Explain photosynthesis in simple words 🌱\n"
        "Example: What are the important points in my uploaded PDF? 📖"
    ),
    height=150,
    key="question_box",
)

send_col, clear_col = st.columns([4, 1])

with send_col:
    ask_button = st.button(
        "🚀 SEARCH / ASK STUDYBUDDY",
        use_container_width=True,
        type="primary",
    )

with clear_col:
    clear_button = st.button("✖ Clear", use_container_width=True)

if clear_button:
    st.session_state.question_box = ""
    st.rerun()

if ask_button:
    current_question = st.session_state.question_box.strip()

    if not current_question:
        st.warning("💡 Please type a question first.")
    elif not ai_configured():
        st.error(
            "🔴 Gemini is not configured. Add a valid GEMINI_API_KEY "
            "to your local secrets or Streamlit Cloud Secrets."
        )
    else:
        answer = run_ai_request(current_question, mode)
        if answer:
            st.markdown("### 🌟 AI Answer")
            render_answer(answer)


# ============================================================
# RECENT HISTORY
# ============================================================
st.markdown("### 🕘 Recent History")

recent_history = get_history(limit=8)

if not recent_history:
    st.info("No questions yet. Ask your first question above! 🚀")
else:
    for row in recent_history:
        star = "⭐" if row["favorite"] else "☆"

        with st.expander(f"{star} {row['question'][:100]}"):
            st.caption(f"{row['created_at']} • {row['mode']}")

            if row["file_name"]:
                st.caption(f"📄 File: {row['file_name']}")

            st.markdown("**Your question**")
            st.write(row["question"])

            st.markdown("**StudyBuddy answer**")
            st.markdown(row["answer"])

            h1, h2 = st.columns(2)

            with h1:
                if st.button(
                    f"{star} Favorite",
                    key=f"fav_{row['id']}",
                    use_container_width=True,
                ):
                    toggle_favorite(row["id"])
                    st.rerun()

            with h2:
                if st.button(
                    "🗑️ Delete",
                    key=f"delete_{row['id']}",
                    use_container_width=True,
                ):
                    delete_chat(row["id"])
                    st.rerun()


# ============================================================
# NOTES
# ============================================================
st.markdown("### 📝 Save a Study Note")

note_title = st.text_input(
    "Note title",
    placeholder="Example: Important Physics Formulas",
)

note_content = st.text_area(
    "Note content",
    value=st.session_state.last_answer,
    placeholder="Write or paste something important to remember...",
    height=130,
)

if st.button("💾 SAVE NOTE", use_container_width=True):
    if not note_title.strip():
        st.warning("Please enter a note title.")
    elif not note_content.strip():
        st.warning("Please enter note content.")
    else:
        save_note(note_title.strip(), note_content.strip())
        st.success("✅ Note saved locally!")
        st.rerun()

saved_notes = get_notes()

if saved_notes:
    st.markdown("### 📚 My Saved Notes")

    for note in saved_notes:
        with st.expander(f"📌 {note['title']}"):
            st.write(note["content"])
            st.caption(f"Saved: {note['created_at']}")

            if st.button(
                "🗑️ Delete Note",
                key=f"note_delete_{note['id']}",
            ):
                delete_note(note["id"])
                st.rerun()


# ============================================================
# FOOTER
# ============================================================
st.markdown(
    """
    <div class="footer">
        <span>🤖 StudyBuddy AI</span>
        <span>Python • Streamlit • Gemini AI</span>
        <span>💜 Learn • Practice • Grow 🚀</span>
    </div>
    """,
    unsafe_allow_html=True,
)
