from pathlib import Path
import sqlite3

from flask import Flask, g, render_template


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

    @app.teardown_appcontext
    def close_db(error=None):
        db = g.pop("db", None)
        if db is not None:
            db.close()

    @app.route("/")
    def dashboard():
        project = get_active_project()
        counts = get_dashboard_counts(project["id"])
        return render_template(
            "dashboard.html",
            title="Dashboard",
            active_page="dashboard",
            project=project,
            counts=counts,
        )

    @app.route("/documents")
    def documents():
        return render_template(
            "documents.html",
            title="Documents",
            active_page="documents",
            panel_title="Documents",
            panel_body=(
                "This section is prepared for later TXT and DOCX import, "
                "document notes, tags, and document-level review."
            ),
        )

    @app.route("/codes")
    def codes():
        return render_template(
            "codes.html",
            title="Codes",
            active_page="codes",
            panel_title="Codes",
            panel_body=(
                "This section is prepared for later open, axial, and category "
                "coding with definitions, inclusion rules, exclusion rules, "
                "examples, colors, and analytical notes."
            ),
        )

    @app.route("/memos")
    def memos():
        return render_template(
            "memos.html",
            title="Memos",
            active_page="memos",
            panel_title="Memos",
            panel_body=(
                "This section is prepared for later research, methodological, "
                "theoretical, and entity-linked memo writing."
            ),
        )

    @app.route("/gt")
    def gt_workspace():
        return render_template(
            "gt_workspace.html",
            title="GT Workspace",
            active_page="gt",
            panel_title="Grounded Theory Workspace",
            panel_body=(
                "This section is prepared for later comparison of incidents, "
                "open codes, axial links, categories, relations, and emerging "
                "theoretical models."
            ),
        )

    @app.route("/cda")
    def cda_workspace():
        return render_template(
            "cda_workspace.html",
            title="CDA Workspace",
            active_page="cda",
            panel_title="Critical Discourse Analysis Workspace",
            panel_body=(
                "This section is prepared for later discourse-analytic helpers "
                "focused on actors, argumentation, framing, modality, ideology, "
                "intertextuality, and power relations."
            ),
        )

    @app.route("/exports")
    def exports():
        return render_template(
            "exports.html",
            title="Exports",
            active_page="exports",
            panel_title="Exports",
            panel_body=(
                "This section is prepared for later export of coded segments, "
                "codebooks, memos, project summaries, and analysis reports."
            ),
        )

    return app


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def init_db_if_needed() -> None:
    is_new_database = not DATABASE.exists()
    db = get_db()

    if is_new_database:
        schema_sql = SCHEMA.read_text(encoding="utf-8")
        db.executescript(schema_sql)
        db.commit()


def get_active_project() -> sqlite3.Row:
    project = get_db().execute(
        "SELECT id, name, description FROM projects ORDER BY id LIMIT 1"
    ).fetchone()
    if project is None:
        raise RuntimeError("The database has no project. Check schema.sql initialization.")
    return project


def get_dashboard_counts(project_id: int) -> dict[str, int]:
    db = get_db()
    return {
        "documents": db.execute(
            "SELECT COUNT(*) FROM documents WHERE project_id = ?", (project_id,)
        ).fetchone()[0],
        "codes": db.execute(
            "SELECT COUNT(*) FROM codes WHERE project_id = ?", (project_id,)
        ).fetchone()[0],
        "segments": db.execute(
            """
            SELECT COUNT(*)
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        ).fetchone()[0],
        "memos": db.execute(
            "SELECT COUNT(*) FROM memos WHERE project_id = ?", (project_id,)
        ).fetchone()[0],
    }


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
