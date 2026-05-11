from datetime import datetime
from pathlib import Path
import sqlite3
from uuid import uuid4

from docx import Document as DocxDocument
from flask import (
    Flask,
    abort,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)
from markupsafe import Markup, escape
from werkzeug.utils import secure_filename


APP_NAME = "discourseLab"
APP_PHASE = "3"
CURRENT_PHASE_LABEL = "Phase 3 — Segment creation"
DEFAULT_PROJECT_NAME = "Demo Project"
DEFAULT_PROJECT_DESCRIPTION = "Initial local discourseLab project."
ALLOWED_EXTENSIONS = {"txt", "docx"}
MAX_UPLOAD_SIZE = 16 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
UPLOAD_DIR = BASE_DIR / "uploads"
EXPORT_DIR = BASE_DIR / "exports"
DATABASE = INSTANCE_DIR / "disLab.sqlite"
SCHEMA = BASE_DIR / "schema.sql"


def create_app() -> Flask:
    app = Flask(__name__, instance_path=str(INSTANCE_DIR))
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE
    app.config["SECRET_KEY"] = "discourseLab-local-development-key"

    INSTANCE_DIR.mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)

    with app.app_context():
        init_db_if_needed()
        get_active_project()

    app.teardown_appcontext(close_db)

    @app.route("/")
    def dashboard():
        active_project = get_active_project()
        counts = get_dashboard_counts(active_project["id"])
        audit_entries = get_latest_audit_entries(active_project["id"])
        latest_documents = get_latest_documents(active_project["id"])
        latest_segments = get_latest_segments(active_project["id"])
        return render_template(
            "dashboard.html",
            title="Dashboard",
            active_page="dashboard",
            active_project=active_project,
            counts=counts,
            audit_entries=audit_entries,
            latest_documents=latest_documents,
            latest_segments=latest_segments,
            current_phase=CURRENT_PHASE_LABEL,
        )

    @app.route("/documents")
    def documents():
        active_project = get_active_project()
        document_rows = get_documents_for_project(active_project["id"])
        return render_template(
            "documents.html",
            title="Documents",
            active_page="documents",
            active_project=active_project,
            documents=document_rows,
        )

    @app.route("/documents/import", methods=["POST"])
    def import_document():
        active_project = get_active_project()
        uploaded_file = request.files.get("document_file")
        title = request.form.get("title", "").strip()
        note = request.form.get("note", "").strip()

        if uploaded_file is None or uploaded_file.filename == "":
            flash("Choose a TXT or DOCX file to import.", "error")
            return redirect(url_for("documents"))

        original_filename = uploaded_file.filename
        extension = get_file_extension(original_filename)
        if extension not in ALLOWED_EXTENSIONS:
            flash("Unsupported file type. Import only TXT or DOCX files.", "error")
            return redirect(url_for("documents"))

        safe_name = build_unique_upload_name(original_filename)
        saved_path = UPLOAD_DIR / safe_name
        uploaded_file.save(saved_path)

        try:
            text_content = extract_text(saved_path, extension)
        except Exception as error:
            saved_path.unlink(missing_ok=True)
            flash(f"Could not extract text from {original_filename}: {error}", "error")
            return redirect(url_for("documents"))

        text_content = normalize_line_endings(text_content).strip()
        if not text_content:
            saved_path.unlink(missing_ok=True)
            flash("The uploaded file did not contain extractable text.", "error")
            return redirect(url_for("documents"))

        if not title:
            title = Path(original_filename).stem

        document_id = execute_write(
            """
            INSERT INTO documents (
                project_id, title, original_filename, file_type, text_content, note
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (active_project["id"], title, original_filename, extension, text_content, note),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="document",
            entity_id=document_id,
            action="import_document",
            details=f"Imported document: {title}",
        )
        flash(f"Imported document: {title}", "success")
        return redirect(url_for("documents"))

    @app.route("/documents/<int:document_id>")
    def document_view(document_id: int):
        active_project = get_active_project()
        document = get_document_for_project(document_id, active_project["id"])
        if document is None:
            flash("Document not found.", "error")
            abort(404)

        segments = get_segments_for_document(document_id)
        highlighted_text = build_highlighted_document_text(
            document["text_content"] or "", segments
        )
        return render_template(
            "document_view.html",
            title=document["title"],
            active_page="documents",
            active_project=active_project,
            document=document,
            text_length=len(document["text_content"] or ""),
            segment_count=len(segments),
            segments=segments,
            highlighted_text=highlighted_text,
        )

    @app.route("/documents/<int:document_id>/segments", methods=["POST"])
    def create_segment(document_id: int):
        active_project = get_active_project()
        document = get_document_for_project(document_id, active_project["id"])
        if document is None:
            flash("Document not found.", "error")
            abort(404)

        segment_name = request.form.get("name", "").strip()
        selected_text = request.form.get("selected_text", "")
        note = request.form.get("note", "").strip()
        try:
            start_offset = int(request.form.get("start_offset", ""))
            end_offset = int(request.form.get("end_offset", ""))
        except ValueError:
            flash("Invalid selection. Select text inside the document panel.", "error")
            return redirect(url_for("document_view", document_id=document_id))

        document_text = document["text_content"] or ""
        if not is_valid_segment_selection(
            document_text, selected_text, start_offset, end_offset
        ):
            flash("Invalid selection. Select a passage inside the document text.", "error")
            return redirect(url_for("document_view", document_id=document_id))

        selected_text = document_text[start_offset:end_offset]
        segment_id = execute_write(
            """
            INSERT INTO segments (
                document_id, name, selected_text, start_offset, end_offset, note
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (document_id, segment_name, selected_text, start_offset, end_offset, note),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="segment",
            entity_id=segment_id,
            action="create_segment",
            details=f"Created segment in document: {document['title']}",
        )
        flash("Segment created.", "success")
        return redirect(url_for("document_view", document_id=document_id))

    @app.route("/documents/<int:document_id>/delete", methods=["POST"])
    def delete_document(document_id: int):
        active_project = get_active_project()
        document = get_document_for_project(document_id, active_project["id"])
        if document is None:
            flash("Document not found.", "error")
            abort(404)

        delete_document_data(document_id)
        log_action(
            project_id=active_project["id"],
            entity_type="document",
            entity_id=document_id,
            action="delete_document",
            details=f"Deleted document: {document['title']}",
        )
        flash(f"Deleted document: {document['title']}", "success")
        return redirect(url_for("documents"))

    @app.route("/segments/<int:segment_id>/delete", methods=["POST"])
    def delete_segment(segment_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Segment not found.", "error")
            abort(404)

        db = get_db()
        db.execute("DELETE FROM segment_codes WHERE segment_id = ?", (segment_id,))
        db.execute("DELETE FROM segments WHERE id = ?", (segment_id,))
        db.commit()
        log_action(
            project_id=active_project["id"],
            entity_type="segment",
            entity_id=segment_id,
            action="delete_segment",
            details=f"Deleted segment from document: {segment['document_title']}",
        )
        flash("Segment deleted.", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/codes")
    def codes():
        return render_placeholder(
            "codes.html",
            title="Codes",
            active_page="codes",
            panel_title="Codes",
            panel_body="This section will manage open, axial, category, and CDA codes in later phases.",
        )

    @app.route("/memos")
    def memos():
        return render_placeholder(
            "memos.html",
            title="Memos",
            active_page="memos",
            panel_title="Memos",
            panel_body=(
                "This section will collect segment, code, methodological, theoretical, "
                "reflexive, and comparison memos."
            ),
        )

    @app.route("/gt")
    def gt_workspace():
        return render_placeholder(
            "gt_workspace.html",
            title="Grounded Theory Workspace",
            active_page="gt",
            panel_title="Grounded Theory Workspace",
            panel_body=(
                "This workspace will support open coding, axial coding, category development, "
                "constant comparison, and theoretical memo-writing."
            ),
        )

    @app.route("/cda")
    def cda_workspace():
        return render_placeholder(
            "cda_workspace.html",
            title="CDA Workspace",
            active_page="cda",
            panel_title="CDA Workspace",
            panel_body=(
                "This workspace will support textual, discursive-practice, and social-practice "
                "analysis, including actors, voice, silence, modality, metaphors, "
                "presuppositions, and legitimation."
            ),
        )

    @app.route("/exports")
    def exports():
        return render_placeholder(
            "exports.html",
            title="Exports",
            active_page="exports",
            panel_title="Exports",
            panel_body=(
                "This section will export codebooks, coded segments, memos, research models, "
                "Markdown, CSV, JSON, PNG/SVG, and LaTeX/TikZ."
            ),
        )

    @app.route("/health")
    def health():
        return jsonify({"status": "ok", "app": APP_NAME, "phase": APP_PHASE})

    @app.errorhandler(404)
    def not_found(error):
        return (
            render_template(
                "error.html",
                title="Not Found",
                active_page="",
                message="The requested discourseLab page or document was not found.",
            ),
            404,
        )

    return app


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(error=None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db_if_needed() -> None:
    is_new_database = not DATABASE.exists()
    db = get_db()
    if is_new_database:
        schema_sql = SCHEMA.read_text(encoding="utf-8")
        db.executescript(schema_sql)
        db.commit()

    run_migrations()


def run_migrations() -> None:
    db = get_db()
    segment_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(segments)").fetchall()
    }
    if "name" not in segment_columns:
        db.execute("ALTER TABLE segments ADD COLUMN name TEXT")
        db.commit()


def query_one(sql: str, params: tuple = ()) -> sqlite3.Row | None:
    return get_db().execute(sql, params).fetchone()


def query_all(sql: str, params: tuple = ()) -> list[sqlite3.Row]:
    return get_db().execute(sql, params).fetchall()


def execute_write(sql: str, params: tuple = ()) -> int:
    cursor = get_db().execute(sql, params)
    get_db().commit()
    return cursor.lastrowid


def get_active_project() -> sqlite3.Row:
    project = query_one(
        """
        SELECT id, name, description, created_at, updated_at
        FROM projects
        ORDER BY id
        LIMIT 1
        """
    )
    if project is not None:
        return project

    project_id = execute_write(
        "INSERT INTO projects (name, description) VALUES (?, ?)",
        (DEFAULT_PROJECT_NAME, DEFAULT_PROJECT_DESCRIPTION),
    )
    log_action(
        project_id=project_id,
        entity_type="project",
        entity_id=project_id,
        action="create_default_project",
        details="Created default discourseLab project.",
    )
    return query_one(
        """
        SELECT id, name, description, created_at, updated_at
        FROM projects
        WHERE id = ?
        """,
        (project_id,),
    )


def get_dashboard_counts(project_id: int) -> dict[str, int]:
    return {
        "documents": query_one(
            "SELECT COUNT(*) AS count FROM documents WHERE project_id = ?", (project_id,)
        )["count"],
        "codes": query_one(
            "SELECT COUNT(*) AS count FROM codes WHERE project_id = ?", (project_id,)
        )["count"],
        "segments": query_one(
            """
            SELECT COUNT(*) AS count
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        )["count"],
        "memos": query_one(
            "SELECT COUNT(*) AS count FROM memos WHERE project_id = ?", (project_id,)
        )["count"],
        "research_questions": query_one(
            "SELECT COUNT(*) AS count FROM research_questions WHERE project_id = ?",
            (project_id,),
        )["count"],
        "relations": query_one(
            "SELECT COUNT(*) AS count FROM relations WHERE project_id = ?", (project_id,)
        )["count"],
    }


def get_documents_for_project(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            documents.id,
            documents.title,
            documents.original_filename,
            documents.file_type,
            documents.created_at,
            LENGTH(COALESCE(documents.text_content, '')) AS text_length,
            COUNT(segments.id) AS segment_count
        FROM documents
        LEFT JOIN segments ON segments.document_id = documents.id
        WHERE documents.project_id = ?
        GROUP BY documents.id
        ORDER BY datetime(documents.created_at) DESC, documents.id DESC
        """,
        (project_id,),
    )


def get_latest_documents(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT title, file_type, created_at
        FROM documents
        WHERE project_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def get_latest_segments(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            COALESCE(segments.name, '') AS name,
            segments.selected_text,
            segments.created_at,
            documents.title AS document_title
        FROM segments
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        ORDER BY datetime(segments.created_at) DESC, segments.id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def get_document_for_project(document_id: int, project_id: int) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT id, project_id, title, original_filename, file_type, text_content,
               note, created_at, updated_at
        FROM documents
        WHERE id = ? AND project_id = ?
        """,
        (document_id, project_id),
    )


def get_document_segment_count(document_id: int) -> int:
    return query_one(
        "SELECT COUNT(*) AS count FROM segments WHERE document_id = ?", (document_id,)
    )["count"]


def get_segments_for_document(document_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT id, document_id, COALESCE(name, '') AS name, selected_text, start_offset, end_offset,
               note, created_at, updated_at
        FROM segments
        WHERE document_id = ?
        ORDER BY start_offset ASC, end_offset ASC, id ASC
        """,
        (document_id,),
    )


def get_segment_for_project(segment_id: int, project_id: int) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT
            segments.id,
            segments.document_id,
            COALESCE(segments.name, '') AS name,
            segments.selected_text,
            documents.title AS document_title
        FROM segments
        JOIN documents ON documents.id = segments.document_id
        WHERE segments.id = ? AND documents.project_id = ?
        """,
        (segment_id, project_id),
    )


def is_valid_segment_selection(
    document_text: str, selected_text: str, start_offset: int, end_offset: int
) -> bool:
    if not selected_text.strip():
        return False
    if start_offset < 0 or end_offset <= start_offset:
        return False
    if end_offset > len(document_text):
        return False
    return bool(document_text[start_offset:end_offset].strip())


def build_highlighted_document_text(
    document_text: str, segments: list[sqlite3.Row]
) -> Markup:
    pieces = []
    cursor = 0

    for segment in segments:
        start = segment["start_offset"]
        end = segment["end_offset"]
        if start < cursor or start < 0 or end > len(document_text) or end <= start:
            continue

        pieces.append(escape(document_text[cursor:start]))
        pieces.append(
            Markup('<mark class="segment-highlight" data-segment-id="{}">{}</mark>').format(
                segment["id"], escape(document_text[start:end])
            )
        )
        cursor = end

    pieces.append(escape(document_text[cursor:]))
    return Markup("").join(pieces)


def delete_document_data(document_id: int) -> None:
    db = get_db()
    segment_ids = [
        row["id"]
        for row in db.execute(
            "SELECT id FROM segments WHERE document_id = ?", (document_id,)
        ).fetchall()
    ]
    if segment_ids:
        placeholders = ",".join("?" for _ in segment_ids)
        db.execute(
            f"DELETE FROM segment_codes WHERE segment_id IN ({placeholders})",
            tuple(segment_ids),
        )
    db.execute("DELETE FROM segments WHERE document_id = ?", (document_id,))
    db.execute("DELETE FROM documents WHERE id = ?", (document_id,))
    db.commit()
    # The original uploaded source file is intentionally left in uploads/ for now.


def log_action(
    project_id: int | None,
    entity_type: str,
    entity_id: int | None,
    action: str,
    details: str,
) -> int:
    return execute_write(
        """
        INSERT INTO audit_log (project_id, entity_type, entity_id, action, details)
        VALUES (?, ?, ?, ?, ?)
        """,
        (project_id, entity_type, entity_id, action, details),
    )


def get_latest_audit_entries(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT entity_type, entity_id, action, details, created_at
        FROM audit_log
        WHERE project_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def render_placeholder(
    template_name: str,
    title: str,
    active_page: str,
    panel_title: str,
    panel_body: str,
) -> str:
    return render_template(
        template_name,
        title=title,
        active_page=active_page,
        active_project=get_active_project(),
        panel_title=panel_title,
        panel_body=panel_body,
    )


def get_file_extension(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip(".")


def build_unique_upload_name(filename: str) -> str:
    safe_name = secure_filename(filename)
    if not safe_name:
        safe_name = f"document-{uuid4().hex}"

    path = UPLOAD_DIR / safe_name
    if not path.exists():
        return safe_name

    stem = Path(safe_name).stem
    suffix = Path(safe_name).suffix
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"{stem}-{timestamp}-{uuid4().hex[:8]}{suffix}"


def extract_text(path: Path, extension: str) -> str:
    if extension == "txt":
        return extract_txt(path)
    if extension == "docx":
        return extract_docx(path)
    raise ValueError("Unsupported file type.")


def extract_txt(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1")


def extract_docx(path: Path) -> str:
    document = DocxDocument(path)
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    return "\n\n".join(paragraphs)


def normalize_line_endings(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
