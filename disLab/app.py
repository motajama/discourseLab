from pathlib import Path
import sqlite3

from flask import Flask, g, jsonify, render_template


APP_NAME = "discourseLab"
APP_PHASE = "1"
DEFAULT_PROJECT_NAME = "Demo Project"
DEFAULT_PROJECT_DESCRIPTION = "Initial local discourseLab project."

BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
DATABASE = INSTANCE_DIR / "disLab.sqlite"
SCHEMA = BASE_DIR / "schema.sql"


def create_app() -> Flask:
    app = Flask(__name__, instance_path=str(INSTANCE_DIR))

    INSTANCE_DIR.mkdir(exist_ok=True)
    (BASE_DIR / "uploads").mkdir(exist_ok=True)
    (BASE_DIR / "exports").mkdir(exist_ok=True)

    with app.app_context():
        init_db_if_needed()
        get_active_project()

    app.teardown_appcontext(close_db)

    @app.route("/")
    def dashboard():
        active_project = get_active_project()
        counts = get_dashboard_counts(active_project["id"])
        audit_entries = get_latest_audit_entries(active_project["id"])
        return render_template(
            "dashboard.html",
            title="Dashboard",
            active_page="dashboard",
            active_project=active_project,
            counts=counts,
            audit_entries=audit_entries,
        )

    @app.route("/documents")
    def documents():
        return render_placeholder(
            "documents.html",
            title="Documents",
            active_page="documents",
            panel_title="Documents",
            panel_body="This section will import and manage TXT/DOCX research documents in Phase 2.",
        )

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
    if DATABASE.exists():
        return

    db = get_db()
    schema_sql = SCHEMA.read_text(encoding="utf-8")
    db.executescript(schema_sql)
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


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
