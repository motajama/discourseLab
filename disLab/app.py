import csv
from datetime import datetime
import io
import json
from pathlib import Path
import re
import sqlite3
import zipfile
from uuid import uuid4

from docx import Document as DocxDocument
from flask import (
    Flask,
    Response,
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
APP_PHASE = "8"
CURRENT_PHASE_LABEL = "Phase 8 — Exports and research outputs"
DEFAULT_PROJECT_NAME = "Demo Project"
DEFAULT_PROJECT_DESCRIPTION = "Initial local discourseLab project."
DEFAULT_CODE_COLOR = "#f4c542"
DEFAULT_CDA_MARKER_COLOR = "#7c9a45"
ALLOWED_EXTENSIONS = {"txt", "docx"}
MAX_UPLOAD_SIZE = 16 * 1024 * 1024
MEMO_TYPES = {
    "project": "Project memo",
    "document": "Document memo",
    "segment": "Segment memo",
    "code": "Code memo",
    "methodological": "Methodological memo",
    "theoretical": "Theoretical memo",
    "reflexive": "Reflexive memo",
    "comparison": "Comparison memo",
    "negative_case": "Negative case memo",
}
MEMO_STATUSES = {
    "draft": "Draft",
    "important": "Important",
    "use_in_article": "Use in article",
    "archived": "Archived",
}
GT_COLUMNS = [
    "gt_conditions",
    "gt_context",
    "gt_actions_interactions",
    "gt_consequences",
    "gt_properties",
    "gt_dimensions",
    "gt_theoretical_note",
]
CDA_MARKER_TYPES = {
    "textual": "Textual",
    "discursive_practice": "Discursive practice",
    "social_practice": "Social practice",
    "actor": "Actor",
    "agency": "Agency",
    "voice": "Voice",
    "silence": "Silence",
    "modality": "Modality",
    "evaluation": "Evaluation",
    "metaphor": "Metaphor",
    "presupposition": "Presupposition",
    "nominalization": "Nominalization",
    "passivization": "Passivization",
    "intertextuality": "Intertextuality",
    "legitimation": "Legitimation",
    "framing": "Framing",
    "ideology": "Ideology",
    "power_relation": "Power relation",
    "other": "Other",
}
ACTOR_TYPES = {
    "individual": "Individual",
    "group": "Group",
    "institution": "Institution",
    "state_actor": "State actor",
    "expert": "Expert",
    "journalist": "Journalist",
    "politician": "Politician",
    "public": "Public",
    "vulnerable_group": "Vulnerable group",
    "abstract_actor": "Abstract actor",
    "other": "Other",
}
ACTOR_RELATION_TYPES = {
    "speaks": "Speaks",
    "is_quoted": "Is quoted",
    "is_spoken_about": "Is spoken about",
    "is_evaluated": "Is evaluated",
    "acts": "Acts",
    "is_acted_upon": "Is acted upon",
    "is_silenced": "Is silenced",
    "is_backgrounded": "Is backgrounded",
    "is_aggregated": "Is aggregated",
    "is_individualized": "Is individualized",
}
DISCOURSE_FEATURE_TYPES = {
    "metaphor": "Metaphor",
    "modality": "Modality",
    "evaluation": "Evaluation",
    "presupposition": "Presupposition",
    "legitimation": "Legitimation",
    "intertextuality": "Intertextuality",
    "framing": "Framing",
    "nominalization": "Nominalization",
    "passivization": "Passivization",
    "agency": "Agency",
    "ideology": "Ideology",
    "power_relation": "Power relation",
    "other": "Other",
}

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
        latest_open_codes = get_latest_open_codes(active_project["id"])
        latest_coded_segments = get_latest_coded_segments(active_project["id"])
        latest_memos = get_latest_memos(active_project["id"])
        codes_missing_definitions = get_codes_missing_definitions(active_project["id"])
        gt_preview = get_gt_structure_preview(active_project["id"])
        cda_preview = get_cda_dashboard_preview(active_project["id"])
        return render_template(
            "dashboard.html",
            title="Dashboard",
            active_page="dashboard",
            active_project=active_project,
            counts=counts,
            audit_entries=audit_entries,
            latest_documents=latest_documents,
            latest_segments=latest_segments,
            latest_open_codes=latest_open_codes,
            latest_coded_segments=latest_coded_segments,
            latest_memos=latest_memos,
            codes_missing_definitions=codes_missing_definitions,
            gt_preview=gt_preview,
            cda_preview=cda_preview,
            current_phase=CURRENT_PHASE_LABEL,
            export_links=get_dashboard_export_links(),
            memo_type_labels=MEMO_TYPES,
            memo_status_labels=MEMO_STATUSES,
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

        segments = get_segments_for_document(document_id, active_project["id"])
        open_codes = get_open_codes_for_project(active_project["id"])
        discourse_markers = get_discourse_markers_for_project(active_project["id"])
        actors = get_actors_for_project(active_project["id"])
        document_memos = get_memos_for_entity(active_project["id"], "document", document_id)
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
            open_codes=open_codes,
            discourse_markers=discourse_markers,
            actors=actors,
            marker_type_labels=CDA_MARKER_TYPES,
            actor_type_labels=ACTOR_TYPES,
            actor_relation_type_labels=ACTOR_RELATION_TYPES,
            feature_type_labels=DISCOURSE_FEATURE_TYPES,
            document_memos=document_memos,
            memo_type_labels=MEMO_TYPES,
            memo_status_labels=MEMO_STATUSES,
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
        db.execute("DELETE FROM segment_discourse_markers WHERE segment_id = ?", (segment_id,))
        db.execute("DELETE FROM segment_actors WHERE segment_id = ?", (segment_id,))
        db.execute("DELETE FROM discourse_features WHERE segment_id = ?", (segment_id,))
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
        active_project = get_active_project()
        code_type_filter = request.args.get("code_type", "all").strip()
        codes_rows = get_codes_for_project(active_project["id"], code_type_filter)
        return render_template(
            "codes.html",
            title="Codes",
            active_page="codes",
            active_project=active_project,
            codes=codes_rows,
            code_type_filter=code_type_filter,
            default_code_color=DEFAULT_CODE_COLOR,
        )

    @app.route("/codes/create", methods=["POST"])
    def create_code():
        active_project = get_active_project()
        name = request.form.get("name", "").strip()
        description = request.form.get("description", "").strip()
        color = normalize_code_color(request.form.get("color", ""))
        document_id = request.form.get("document_id", "").strip()

        if not name:
            flash("Code name is required.", "error")
            return redirect_after_code_change(document_id)

        code_id = execute_write(
            """
            INSERT INTO codes (project_id, name, description, code_type, color)
            VALUES (?, ?, ?, 'open', ?)
            """,
            (active_project["id"], name, description, color),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="code",
            entity_id=code_id,
            action="create_open_code",
            details=f"Created open code: {name}",
        )
        flash(f"Created open code: {name}", "success")
        return redirect_after_code_change(document_id)

    @app.route("/codes/<int:code_id>")
    def code_detail(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Code not found.", "error")
            abort(404)

        return render_template(
            "code_detail.html",
            title=f"Code: {code['name']}",
            active_page="codes",
            active_project=active_project,
            code=code,
            usage_count=get_code_usage_count(code_id),
            code_memos=get_memos_for_entity(active_project["id"], "code", code_id),
            coded_segments=get_segments_for_code(code_id, active_project["id"]),
            parent_code=get_parent_code(code["parent_id"], active_project["id"]) if code["parent_id"] else None,
            child_open_codes=get_child_codes(code_id, active_project["id"], "open"),
            child_axial_codes=get_child_codes(code_id, active_project["id"], "axial"),
            category_open_codes=get_open_codes_under_category(code_id, active_project["id"]),
            hierarchy_segments=get_hierarchy_segments_for_code(code, active_project["id"]),
            memo_type_labels=MEMO_TYPES,
            memo_status_labels=MEMO_STATUSES,
        )

    @app.route("/codes/<int:code_id>/edit")
    def edit_code(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Invalid code.", "error")
            abort(404)

        return render_template(
            "code_edit.html",
            title=f"Edit Code: {code['name']}",
            active_page="codes",
            active_project=active_project,
            code=code,
            axial_codes=get_gt_axial_codes(active_project["id"]),
            categories=get_gt_categories(active_project["id"]),
        )

    @app.route("/codes/<int:code_id>/edit", methods=["POST"])
    def update_code(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Invalid code.", "error")
            abort(404)

        name = request.form.get("name", "").strip()
        description = request.form.get("description", "").strip()
        color = normalize_code_color(request.form.get("color", ""))
        parent_id_raw = request.form.get("parent_id", "").strip()
        definition = request.form.get("definition", "").strip()
        include_when = request.form.get("include_when", "").strip()
        exclude_when = request.form.get("exclude_when", "").strip()
        example = request.form.get("example", "").strip()
        analytical_note = request.form.get("analytical_note", "").strip()
        if not name:
            flash("Code name is required.", "error")
            return redirect(url_for("edit_code", code_id=code_id))
        parent_id = None
        if parent_id_raw:
            if not parent_id_raw.isdigit():
                flash("Invalid GT hierarchy operation.", "error")
                return redirect(url_for("edit_code", code_id=code_id))
            parent_id = int(parent_id_raw)
            valid, message = validate_parent_assignment(
                code_id, code["code_type"], parent_id, active_project["id"]
            )
            if not valid:
                flash(message, "error")
                return redirect(url_for("edit_code", code_id=code_id))

        execute_write(
            """
            UPDATE codes
            SET name = ?, description = ?, color = ?, definition = ?,
                include_when = ?, exclude_when = ?, example = ?,
                analytical_note = ?, parent_id = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ? AND project_id = ?
            """,
            (
                name,
                description,
                color,
                definition,
                include_when,
                exclude_when,
                example,
                analytical_note,
                parent_id,
                code_id,
                active_project["id"],
            ),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="code",
            entity_id=code_id,
            action="update_open_code",
            details=f"Updated open code: {name}",
        )
        flash(f"Updated codebook entry: {name}", "success")
        return redirect(url_for("codes"))

    @app.route("/codes/<int:code_id>/delete", methods=["POST"])
    def delete_code(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Invalid code.", "error")
            abort(404)

        db = get_db()
        if code["code_type"] == "open":
            db.execute("DELETE FROM segment_codes WHERE code_id = ?", (code_id,))
        if code["code_type"] == "axial":
            db.execute(
                "UPDATE codes SET parent_id = NULL WHERE parent_id = ? AND code_type = 'open'",
                (code_id,),
            )
        if code["code_type"] == "category":
            db.execute(
                "UPDATE codes SET parent_id = NULL WHERE parent_id = ? AND code_type = 'axial'",
                (code_id,),
            )
        db.execute(
            "DELETE FROM codes WHERE id = ? AND project_id = ?",
            (code_id, active_project["id"]),
        )
        db.commit()
        log_action(
            project_id=active_project["id"],
            entity_type="code",
            entity_id=code_id,
            action="delete_open_code",
            details=f"Deleted {code['code_type']} code: {code['name']}",
        )
        flash(f"Deleted code: {code['name']}", "success")
        return redirect(url_for("codes"))

    @app.route("/segments/<int:segment_id>/codes", methods=["POST"])
    def assign_code_to_segment(segment_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)

        try:
            code_id = int(request.form.get("code_id", ""))
        except ValueError:
            flash("Invalid code.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))

        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Invalid code.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))

        if segment_has_code(segment_id, code_id):
            flash("Code already assigned to this segment.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))

        execute_write(
            "INSERT INTO segment_codes (segment_id, code_id) VALUES (?, ?)",
            (segment_id, code_id),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            project_id=active_project["id"],
            entity_type="segment",
            entity_id=segment_id,
            action="assign_open_code",
            details=f"Assigned code {code['name']} to {segment_label}",
        )
        flash(f"Assigned code: {code['name']}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/segments/<int:segment_id>/codes/<int:code_id>/remove", methods=["POST"])
    def remove_code_from_segment(segment_id: int, code_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)

        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Invalid code.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))

        execute_write(
            "DELETE FROM segment_codes WHERE segment_id = ? AND code_id = ?",
            (segment_id, code_id),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            project_id=active_project["id"],
            entity_type="segment",
            entity_id=segment_id,
            action="remove_open_code",
            details=f"Removed code {code['name']} from {segment_label}",
        )
        flash(f"Removed code: {code['name']}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/memos")
    def memos():
        active_project = get_active_project()
        filters = {
            "memo_type": request.args.get("memo_type", "").strip(),
            "status": request.args.get("status", "").strip(),
            "linked_entity_type": request.args.get("linked_entity_type", "").strip(),
            "linked_entity_id": request.args.get("linked_entity_id", "").strip(),
        }
        memo_rows = get_memos_for_project(active_project["id"], filters)
        return render_template(
            "memos.html",
            title="Memos",
            active_page="memos",
            active_project=active_project,
            memos=memo_rows,
            filters=filters,
            memo_type_labels=MEMO_TYPES,
            memo_status_labels=MEMO_STATUSES,
            documents=get_documents_for_memo_links(active_project["id"]),
            segments=get_segments_for_memo_links(active_project["id"]),
            codes=get_open_codes_for_project(active_project["id"]),
        )

    @app.route("/memos/create", methods=["POST"])
    def create_memo():
        active_project = get_active_project()
        data, error = validate_memo_form(active_project["id"])
        next_url = safe_next_url(request.form.get("next_url", ""))
        if error:
            flash(error, "error")
            return redirect(next_url or url_for("memos"))

        memo_id = execute_write(
            """
            INSERT INTO memos (
                project_id, title, body, memo_type, linked_entity_type,
                linked_entity_id, status
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                active_project["id"],
                data["title"],
                data["body"],
                data["memo_type"],
                data["linked_entity_type"],
                data["linked_entity_id"],
                data["status"],
            ),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="memo",
            entity_id=memo_id,
            action="create_memo",
            details=f"Created memo: {data['title']}",
        )
        flash(f"Created memo: {data['title']}", "success")
        return redirect(next_url or url_for("memos"))

    @app.route("/memos/<int:memo_id>/edit")
    def edit_memo(memo_id: int):
        active_project = get_active_project()
        memo = get_memo_for_project(memo_id, active_project["id"])
        if memo is None:
            flash("Memo not found.", "error")
            abort(404)

        return render_template(
            "memo_edit.html",
            title=f"Edit Memo: {memo['title']}",
            active_page="memos",
            active_project=active_project,
            memo=memo,
            memo_type_labels=MEMO_TYPES,
            memo_status_labels=MEMO_STATUSES,
            documents=get_documents_for_memo_links(active_project["id"]),
            segments=get_segments_for_memo_links(active_project["id"]),
            codes=get_open_codes_for_project(active_project["id"]),
        )

    @app.route("/memos/<int:memo_id>/edit", methods=["POST"])
    def update_memo(memo_id: int):
        active_project = get_active_project()
        memo = get_memo_for_project(memo_id, active_project["id"])
        if memo is None:
            flash("Memo not found.", "error")
            abort(404)

        data, error = validate_memo_form(active_project["id"])
        if error:
            flash(error, "error")
            return redirect(url_for("edit_memo", memo_id=memo_id))

        execute_write(
            """
            UPDATE memos
            SET title = ?, body = ?, memo_type = ?, linked_entity_type = ?,
                linked_entity_id = ?, status = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ? AND project_id = ?
            """,
            (
                data["title"],
                data["body"],
                data["memo_type"],
                data["linked_entity_type"],
                data["linked_entity_id"],
                data["status"],
                memo_id,
                active_project["id"],
            ),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="memo",
            entity_id=memo_id,
            action="update_memo",
            details=f"Updated memo: {data['title']}",
        )
        flash(f"Updated memo: {data['title']}", "success")
        return redirect(url_for("memos"))

    @app.route("/memos/<int:memo_id>/delete", methods=["POST"])
    def delete_memo(memo_id: int):
        active_project = get_active_project()
        memo = get_memo_for_project(memo_id, active_project["id"])
        if memo is None:
            flash("Memo not found.", "error")
            abort(404)

        execute_write(
            "DELETE FROM memos WHERE id = ? AND project_id = ?",
            (memo_id, active_project["id"]),
        )
        log_action(
            project_id=active_project["id"],
            entity_type="memo",
            entity_id=memo_id,
            action="delete_memo",
            details=f"Deleted memo: {memo['title']}",
        )
        flash(f"Deleted memo: {memo['title']}", "success")
        return redirect(url_for("memos"))

    @app.route("/gt")
    def gt_workspace():
        active_project = get_active_project()
        return render_template(
            "gt_workspace.html",
            title="Grounded Theory Workspace",
            active_page="gt",
            active_project=active_project,
            open_codes=get_gt_open_codes(active_project["id"]),
            axial_codes=get_gt_axial_codes(active_project["id"]),
            categories=get_gt_categories(active_project["id"]),
            default_code_color=DEFAULT_CODE_COLOR,
        )

    @app.route("/gt/axial/create", methods=["POST"])
    def create_axial_code():
        return create_gt_code("axial", "create_axial_code", "Created axial code")

    @app.route("/gt/category/create", methods=["POST"])
    def create_category_code():
        return create_gt_code("category", "create_category", "Created category")

    @app.route("/gt/open/<int:open_code_id>/assign-axial", methods=["POST"])
    def assign_open_to_axial(open_code_id: int):
        return assign_code_parent(
            child_id=open_code_id,
            child_type="open",
            parent_type="axial",
            form_field="axial_code_id",
            action="assign_open_to_axial",
            success_template="Assigned open code {child} to axial code {parent}",
        )

    @app.route("/gt/open/<int:open_code_id>/unassign", methods=["POST"])
    def unassign_open_from_axial(open_code_id: int):
        return unassign_code_parent(open_code_id, "open", "unassign_open_from_axial")

    @app.route("/gt/axial/<int:axial_code_id>/assign-category", methods=["POST"])
    def assign_axial_to_category(axial_code_id: int):
        return assign_code_parent(
            child_id=axial_code_id,
            child_type="axial",
            parent_type="category",
            form_field="category_code_id",
            action="assign_axial_to_category",
            success_template="Assigned axial code {child} to category {parent}",
        )

    @app.route("/gt/axial/<int:axial_code_id>/unassign", methods=["POST"])
    def unassign_axial_from_category(axial_code_id: int):
        return unassign_code_parent(axial_code_id, "axial", "unassign_axial_from_category")

    @app.route("/gt/codes/<int:code_id>/edit")
    def edit_gt_code(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None:
            flash("Code not found.", "error")
            abort(404)
        if code["code_type"] == "open":
            return redirect(url_for("edit_code", code_id=code_id))
        return render_template(
            "gt_code_edit.html",
            title=f"Edit GT Code: {code['name']}",
            active_page="gt",
            active_project=active_project,
            code=code,
        )

    @app.route("/gt/codes/<int:code_id>/edit", methods=["POST"])
    def update_gt_code(code_id: int):
        active_project = get_active_project()
        code = get_code_for_project(code_id, active_project["id"])
        if code is None or code["code_type"] not in {"axial", "category"}:
            flash("Code not found.", "error")
            abort(404)

        name = request.form.get("name", "").strip()
        if not name:
            flash("Code name is required.", "error")
            return redirect(url_for("edit_gt_code", code_id=code_id))
        values = (
            name,
            request.form.get("description", "").strip(),
            normalize_code_color(request.form.get("color", "")),
            request.form.get("definition", "").strip(),
            request.form.get("analytical_note", "").strip(),
            request.form.get("gt_conditions", "").strip(),
            request.form.get("gt_context", "").strip(),
            request.form.get("gt_actions_interactions", "").strip(),
            request.form.get("gt_consequences", "").strip(),
            request.form.get("gt_properties", "").strip(),
            request.form.get("gt_dimensions", "").strip(),
            request.form.get("gt_theoretical_note", "").strip(),
            code_id,
            active_project["id"],
        )
        execute_write(
            """
            UPDATE codes
            SET name = ?, description = ?, color = ?, definition = ?,
                analytical_note = ?, gt_conditions = ?, gt_context = ?,
                gt_actions_interactions = ?, gt_consequences = ?,
                gt_properties = ?, gt_dimensions = ?, gt_theoretical_note = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ? AND project_id = ?
            """,
            values,
        )
        log_action(
            active_project["id"],
            "code",
            code_id,
            "update_gt_code",
            f"Updated GT code: {name}",
        )
        flash(f"Updated GT code: {name}", "success")
        return redirect(url_for("gt_workspace"))

    @app.route("/gt/compare")
    def gt_compare():
        active_project = get_active_project()
        open_codes = get_open_codes_for_project(active_project["id"])
        code_a = get_compare_code(request.args.get("code_a"), active_project["id"])
        code_b = get_compare_code(request.args.get("code_b"), active_project["id"])
        return render_template(
            "gt_compare.html",
            title="Constant Comparison",
            active_page="gt",
            active_project=active_project,
            open_codes=open_codes,
            code_a=code_a,
            code_b=code_b,
            memo_status_labels=MEMO_STATUSES,
        )

    @app.route("/cda")
    def cda_workspace():
        active_project = get_active_project()
        return render_template(
            "cda_workspace.html",
            title="CDA Workspace",
            active_page="cda",
            active_project=active_project,
            counts=get_cda_counts(active_project["id"]),
            markers=get_discourse_markers_for_project(active_project["id"]),
            actors=get_actors_for_project(active_project["id"]),
            marker_type_labels=CDA_MARKER_TYPES,
            actor_type_labels=ACTOR_TYPES,
            default_marker_color=DEFAULT_CDA_MARKER_COLOR,
        )

    @app.route("/cda/markers/create", methods=["POST"])
    def create_discourse_marker():
        active_project = get_active_project()
        name = request.form.get("name", "").strip()
        marker_type = request.form.get("marker_type", "").strip()
        description = request.form.get("description", "").strip()
        color = normalize_cda_color(request.form.get("color", ""))
        if not name:
            flash("CDA marker name is required.", "error")
            return redirect(url_for("cda_workspace"))
        if marker_type not in CDA_MARKER_TYPES:
            flash("Invalid marker.", "error")
            return redirect(url_for("cda_workspace"))
        marker_id = execute_write(
            """
            INSERT INTO discourse_markers (project_id, name, marker_type, description, color)
            VALUES (?, ?, ?, ?, ?)
            """,
            (active_project["id"], name, marker_type, description, color),
        )
        log_action(
            active_project["id"],
            "discourse_marker",
            marker_id,
            "create_discourse_marker",
            f"Created CDA marker: {name}",
        )
        flash(f"CDA marker created: {name}", "success")
        return redirect(url_for("cda_workspace"))

    @app.route("/cda/markers/<int:marker_id>/delete", methods=["POST"])
    def delete_discourse_marker(marker_id: int):
        active_project = get_active_project()
        marker = get_discourse_marker_for_project(marker_id, active_project["id"])
        if marker is None:
            flash("Invalid marker.", "error")
            return redirect(url_for("cda_workspace"))
        db = get_db()
        db.execute("DELETE FROM segment_discourse_markers WHERE marker_id = ?", (marker_id,))
        db.execute(
            "DELETE FROM discourse_markers WHERE id = ? AND project_id = ?",
            (marker_id, active_project["id"]),
        )
        db.commit()
        log_action(
            active_project["id"],
            "discourse_marker",
            marker_id,
            "delete_discourse_marker",
            f"Deleted CDA marker: {marker['name']}",
        )
        flash(f"CDA marker deleted: {marker['name']}", "success")
        return redirect(url_for("cda_workspace"))

    @app.route("/cda/actors/create", methods=["POST"])
    def create_actor():
        active_project = get_active_project()
        name = request.form.get("name", "").strip()
        actor_type = request.form.get("actor_type", "").strip()
        description = request.form.get("description", "").strip()
        if not name:
            flash("Actor name is required.", "error")
            return redirect(url_for("cda_workspace"))
        if actor_type not in ACTOR_TYPES:
            flash("Invalid actor.", "error")
            return redirect(url_for("cda_workspace"))
        actor_id = execute_write(
            """
            INSERT INTO actors (project_id, name, actor_type, description)
            VALUES (?, ?, ?, ?)
            """,
            (active_project["id"], name, actor_type, description),
        )
        log_action(
            active_project["id"],
            "actor",
            actor_id,
            "create_actor",
            f"Created actor: {name}",
        )
        flash(f"Actor created: {name}", "success")
        return redirect(url_for("cda_workspace"))

    @app.route("/cda/actors/<int:actor_id>/delete", methods=["POST"])
    def delete_actor(actor_id: int):
        active_project = get_active_project()
        actor = get_actor_for_project(actor_id, active_project["id"])
        if actor is None:
            flash("Invalid actor.", "error")
            return redirect(url_for("cda_workspace"))
        db = get_db()
        db.execute("DELETE FROM segment_actors WHERE actor_id = ?", (actor_id,))
        db.execute(
            "DELETE FROM actors WHERE id = ? AND project_id = ?",
            (actor_id, active_project["id"]),
        )
        db.commit()
        log_action(
            active_project["id"],
            "actor",
            actor_id,
            "delete_actor",
            f"Deleted actor: {actor['name']}",
        )
        flash(f"Actor deleted: {actor['name']}", "success")
        return redirect(url_for("cda_workspace"))

    @app.route("/segments/<int:segment_id>/discourse-markers", methods=["POST"])
    def assign_discourse_marker_to_segment(segment_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)
        try:
            marker_id = int(request.form.get("marker_id", ""))
        except ValueError:
            flash("Invalid marker.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        marker = get_discourse_marker_for_project(marker_id, active_project["id"])
        if marker is None:
            flash("Invalid marker.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        if segment_has_discourse_marker(segment_id, marker_id):
            flash("CDA marker already assigned to this segment.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        execute_write(
            """
            INSERT INTO segment_discourse_markers (segment_id, marker_id, note)
            VALUES (?, ?, ?)
            """,
            (segment_id, marker_id, request.form.get("note", "").strip()),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "assign_discourse_marker",
            f"Assigned CDA marker {marker['name']} to {segment_label}",
        )
        flash(f"Marker assigned to segment: {marker['name']}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/segments/<int:segment_id>/discourse-markers/<int:marker_id>/remove", methods=["POST"])
    def remove_discourse_marker_from_segment(segment_id: int, marker_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        marker = get_discourse_marker_for_project(marker_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)
        if marker is None:
            flash("Invalid marker.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        execute_write(
            "DELETE FROM segment_discourse_markers WHERE segment_id = ? AND marker_id = ?",
            (segment_id, marker_id),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "remove_discourse_marker",
            f"Removed CDA marker {marker['name']} from {segment_label}",
        )
        flash(f"Marker removed from segment: {marker['name']}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/segments/<int:segment_id>/actors", methods=["POST"])
    def assign_actor_to_segment(segment_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)
        try:
            actor_id = int(request.form.get("actor_id", ""))
        except ValueError:
            flash("Invalid actor.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        relation_type = request.form.get("relation_type", "").strip()
        actor = get_actor_for_project(actor_id, active_project["id"])
        if actor is None:
            flash("Invalid actor.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        if relation_type not in ACTOR_RELATION_TYPES:
            flash("Invalid actor relation.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        execute_write(
            """
            INSERT INTO segment_actors (segment_id, actor_id, relation_type, note)
            VALUES (?, ?, ?, ?)
            """,
            (segment_id, actor_id, relation_type, request.form.get("note", "").strip()),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "assign_actor_to_segment",
            f"Assigned actor {actor['name']} as {relation_type} to {segment_label}",
        )
        flash(f"Actor assigned to segment: {actor['name']}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/segments/<int:segment_id>/actors/<int:segment_actor_id>/remove", methods=["POST"])
    def remove_actor_from_segment(segment_id: int, segment_actor_id: int):
        active_project = get_active_project()
        segment_actor = get_segment_actor_for_project(segment_actor_id, segment_id, active_project["id"])
        if segment_actor is None:
            flash("Invalid actor.", "error")
            abort(404)
        segment_label = segment_actor["segment_name"] or f"segment {segment_id}"
        execute_write("DELETE FROM segment_actors WHERE id = ?", (segment_actor_id,))
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "remove_actor_from_segment",
            f"Removed actor annotation from {segment_label}",
        )
        flash("Actor annotation removed from segment.", "success")
        return redirect(url_for("document_view", document_id=segment_actor["document_id"]))

    @app.route("/segments/<int:segment_id>/features", methods=["POST"])
    def create_discourse_feature(segment_id: int):
        active_project = get_active_project()
        segment = get_segment_for_project(segment_id, active_project["id"])
        if segment is None:
            flash("Invalid segment.", "error")
            abort(404)
        feature_type = request.form.get("feature_type", "").strip()
        value = request.form.get("value", "").strip()
        interpretation = request.form.get("interpretation", "").strip()
        if feature_type not in DISCOURSE_FEATURE_TYPES:
            flash("Invalid feature.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        if not value:
            flash("Invalid feature.", "error")
            return redirect(url_for("document_view", document_id=segment["document_id"]))
        feature_id = execute_write(
            """
            INSERT INTO discourse_features (segment_id, feature_type, value, interpretation)
            VALUES (?, ?, ?, ?)
            """,
            (segment_id, feature_type, value, interpretation),
        )
        segment_label = segment["name"] or f"segment {segment_id}"
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "create_discourse_feature",
            f"Created discourse feature {feature_type} for {segment_label}",
        )
        flash(f"Discourse feature created: {feature_type}", "success")
        return redirect(url_for("document_view", document_id=segment["document_id"]))

    @app.route("/segments/<int:segment_id>/features/<int:feature_id>/delete", methods=["POST"])
    def delete_discourse_feature(segment_id: int, feature_id: int):
        active_project = get_active_project()
        feature = get_discourse_feature_for_project(feature_id, segment_id, active_project["id"])
        if feature is None:
            flash("Invalid feature.", "error")
            abort(404)
        segment_label = feature["segment_name"] or f"segment {segment_id}"
        execute_write("DELETE FROM discourse_features WHERE id = ?", (feature_id,))
        log_action(
            active_project["id"],
            "segment",
            segment_id,
            "delete_discourse_feature",
            f"Deleted discourse feature {feature['feature_type']} from {segment_label}",
        )
        flash(f"Discourse feature deleted: {feature['feature_type']}", "success")
        return redirect(url_for("document_view", document_id=feature["document_id"]))

    @app.route("/cda/features")
    def cda_features():
        active_project = get_active_project()
        filters = {
            "feature_type": request.args.get("feature_type", "").strip(),
            "document_id": request.args.get("document_id", "").strip(),
        }
        return render_template(
            "cda_features.html",
            title="CDA Feature Overview",
            active_page="cda",
            active_project=active_project,
            features=get_discourse_features_for_project(active_project["id"], filters),
            counts=get_discourse_feature_counts(active_project["id"]),
            documents=get_documents_for_memo_links(active_project["id"]),
            filters=filters,
            feature_type_labels=DISCOURSE_FEATURE_TYPES,
        )

    @app.route("/cda/voice-silence")
    def cda_voice_silence():
        active_project = get_active_project()
        return render_template(
            "cda_voice_silence.html",
            title="Voice and Silence Report",
            active_page="cda",
            active_project=active_project,
            rows=get_voice_silence_report(active_project["id"]),
            relation_type_labels=ACTOR_RELATION_TYPES,
        )

    @app.route("/exports")
    def exports():
        export_sections = [
            {
                "title": "Codebook exports",
                "cards": [
                    {
                        "title": "Codebook Markdown",
                        "description": "Complete codebook with hierarchy, GT fields, usage counts, and memo counts.",
                        "format": "Markdown",
                        "endpoint": "export_codebook_markdown",
                        "button": "Download codebook",
                    },
                ],
            },
            {
                "title": "Segment exports",
                "cards": [
                    {
                        "title": "Coded segments CSV",
                        "description": "One CSV row per segment-code assignment.",
                        "format": "CSV",
                        "endpoint": "export_coded_segments_csv",
                        "button": "Download CSV",
                    },
                    {
                        "title": "Coded segments Markdown",
                        "description": "Readable coded segment report grouped by document.",
                        "format": "Markdown",
                        "endpoint": "export_coded_segments_markdown",
                        "button": "Download Markdown",
                    },
                ],
            },
            {
                "title": "Memo exports",
                "cards": [
                    {
                        "title": "Memos Markdown",
                        "description": "All project memos grouped by memo type and status.",
                        "format": "Markdown",
                        "endpoint": "export_memos_markdown",
                        "button": "Download memos",
                    },
                ],
            },
            {
                "title": "Grounded Theory exports",
                "cards": [
                    {
                        "title": "GT hierarchy Markdown",
                        "description": "Categories, axial codes, open codes, and representative segments.",
                        "format": "Markdown",
                        "endpoint": "export_gt_hierarchy_markdown",
                        "button": "Download GT hierarchy",
                    },
                ],
            },
            {
                "title": "CDA exports",
                "cards": [
                    {
                        "title": "CDA features CSV",
                        "description": "All discourse features with document and segment context.",
                        "format": "CSV",
                        "endpoint": "export_cda_features_csv",
                        "button": "Download features",
                    },
                    {
                        "title": "Actor voice/silence CSV",
                        "description": "Actor-level voice, silence, evaluation, and agency counts.",
                        "format": "CSV",
                        "endpoint": "export_voice_silence_csv",
                        "button": "Download report",
                    },
                ],
            },
            {
                "title": "Project exports",
                "cards": [
                    {
                        "title": "Project summary Markdown",
                        "description": "Readable project overview with counts, top items, and audit summary.",
                        "format": "Markdown",
                        "endpoint": "export_project_summary_markdown",
                        "button": "Download summary",
                    },
                    {
                        "title": "Full project JSON",
                        "description": "Structured active-project data for archiving and inspection.",
                        "format": "JSON",
                        "endpoint": "export_project_json",
                        "button": "Download JSON",
                    },
                    {
                        "title": "Complete research package ZIP",
                        "description": "All Phase 8 exports bundled into one ZIP file.",
                        "format": "ZIP",
                        "endpoint": "export_project_package",
                        "button": "Download package",
                    },
                ],
            },
        ]
        return render_template(
            "exports.html",
            title="Exports",
            active_page="exports",
            active_project=get_active_project(),
            export_sections=export_sections,
        )

    @app.route("/exports/codebook.md")
    def export_codebook_markdown():
        active_project = get_active_project()
        return download_text(
            build_codebook_markdown(active_project),
            "discourseLab_codebook.md",
            "text/markdown; charset=utf-8",
        )

    @app.route("/exports/coded-segments.csv")
    def export_coded_segments_csv():
        active_project = get_active_project()
        return download_text(
            generate_coded_segments_csv(active_project),
            "discourseLab_coded_segments.csv",
            "text/csv; charset=utf-8",
        )

    @app.route("/exports/coded-segments.md")
    def export_coded_segments_markdown():
        active_project = get_active_project()
        return download_text(
            generate_coded_segments_markdown(active_project),
            "discourseLab_coded_segments.md",
            "text/markdown; charset=utf-8",
        )

    @app.route("/exports/memos.md")
    def export_memos_markdown():
        active_project = get_active_project()
        return download_text(
            generate_memos_markdown(active_project),
            "discourseLab_memos.md",
            "text/markdown; charset=utf-8",
        )

    @app.route("/exports/gt-hierarchy.md")
    def export_gt_hierarchy_markdown():
        active_project = get_active_project()
        return download_text(
            generate_gt_hierarchy_markdown(active_project),
            "discourseLab_gt_hierarchy.md",
            "text/markdown; charset=utf-8",
        )

    @app.route("/exports/cda-features.csv")
    def export_cda_features_csv():
        active_project = get_active_project()
        return download_text(
            generate_cda_features_csv(active_project),
            "discourseLab_cda_features.csv",
            "text/csv; charset=utf-8",
        )

    @app.route("/exports/voice-silence.csv")
    def export_voice_silence_csv():
        active_project = get_active_project()
        return download_text(
            generate_voice_silence_csv(active_project),
            "discourseLab_voice_silence.csv",
            "text/csv; charset=utf-8",
        )

    @app.route("/exports/project-summary.md")
    def export_project_summary_markdown():
        active_project = get_active_project()
        return download_text(
            generate_project_summary_markdown(active_project),
            "discourseLab_project_summary.md",
            "text/markdown; charset=utf-8",
        )

    @app.route("/exports/project.json")
    def export_project_json():
        active_project = get_active_project()
        return download_text(
            generate_project_json(active_project),
            "discourseLab_project.json",
            "application/json; charset=utf-8",
        )

    @app.route("/exports/package.zip")
    def export_project_package():
        active_project = get_active_project()
        return download_binary(
            generate_project_package_zip(active_project),
            "discourseLab_research_package.zip",
            "application/zip",
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
    code_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(codes)").fetchall()
    }
    for column in GT_COLUMNS:
        if column not in code_columns:
            db.execute(f"ALTER TABLE codes ADD COLUMN {column} TEXT")
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS discourse_markers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            marker_type TEXT NOT NULL,
            description TEXT,
            color TEXT NOT NULL DEFAULT '#7c9a45',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (project_id) REFERENCES projects (id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS segment_discourse_markers (
            segment_id INTEGER NOT NULL,
            marker_id INTEGER NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (segment_id, marker_id),
            FOREIGN KEY (segment_id) REFERENCES segments (id) ON DELETE CASCADE,
            FOREIGN KEY (marker_id) REFERENCES discourse_markers (id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS actors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            actor_type TEXT NOT NULL,
            description TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (project_id) REFERENCES projects (id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS segment_actors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            segment_id INTEGER NOT NULL,
            actor_id INTEGER NOT NULL,
            relation_type TEXT NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (segment_id) REFERENCES segments (id) ON DELETE CASCADE,
            FOREIGN KEY (actor_id) REFERENCES actors (id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS discourse_features (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            segment_id INTEGER NOT NULL,
            feature_type TEXT NOT NULL,
            value TEXT NOT NULL,
            interpretation TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (segment_id) REFERENCES segments (id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_discourse_markers_project_id
            ON discourse_markers (project_id);
        CREATE INDEX IF NOT EXISTS idx_segment_discourse_markers_segment_id
            ON segment_discourse_markers (segment_id);
        CREATE INDEX IF NOT EXISTS idx_segment_discourse_markers_marker_id
            ON segment_discourse_markers (marker_id);
        CREATE INDEX IF NOT EXISTS idx_actors_project_id ON actors (project_id);
        CREATE INDEX IF NOT EXISTS idx_segment_actors_segment_id
            ON segment_actors (segment_id);
        CREATE INDEX IF NOT EXISTS idx_segment_actors_actor_id
            ON segment_actors (actor_id);
        CREATE INDEX IF NOT EXISTS idx_discourse_features_segment_id
            ON discourse_features (segment_id);
        """
    )
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
        "open_codes": query_one(
            """
            SELECT COUNT(*) AS count
            FROM codes
            WHERE project_id = ? AND code_type = 'open'
            """,
            (project_id,),
        )["count"],
        "axial_codes": query_one(
            "SELECT COUNT(*) AS count FROM codes WHERE project_id = ? AND code_type = 'axial'",
            (project_id,),
        )["count"],
        "categories": query_one(
            "SELECT COUNT(*) AS count FROM codes WHERE project_id = ? AND code_type = 'category'",
            (project_id,),
        )["count"],
        "cda_markers": query_one(
            "SELECT COUNT(*) AS count FROM discourse_markers WHERE project_id = ?",
            (project_id,),
        )["count"],
        "actors": query_one(
            "SELECT COUNT(*) AS count FROM actors WHERE project_id = ?", (project_id,)
        )["count"],
        "discourse_features": query_one(
            """
            SELECT COUNT(*) AS count
            FROM discourse_features
            JOIN segments ON segments.id = discourse_features.segment_id
            JOIN documents ON documents.id = segments.document_id
            WHERE documents.project_id = ?
            """,
            (project_id,),
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
        "coded_segments": query_one(
            """
            SELECT COUNT(DISTINCT segments.id) AS count
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            JOIN segment_codes ON segment_codes.segment_id = segments.id
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


def get_cda_dashboard_preview(project_id: int) -> dict[str, str | int]:
    relation = query_one(
        """
        SELECT segment_actors.relation_type, COUNT(*) AS count
        FROM segment_actors
        JOIN segments ON segments.id = segment_actors.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        GROUP BY segment_actors.relation_type
        ORDER BY count DESC, segment_actors.relation_type
        LIMIT 1
        """,
        (project_id,),
    )
    feature = query_one(
        """
        SELECT discourse_features.feature_type, COUNT(*) AS count
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        GROUP BY discourse_features.feature_type
        ORDER BY count DESC, discourse_features.feature_type
        LIMIT 1
        """,
        (project_id,),
    )
    return {
        "segments_with_markers": query_one(
            """
            SELECT COUNT(DISTINCT segments.id) AS count
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            JOIN segment_discourse_markers
                ON segment_discourse_markers.segment_id = segments.id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        )["count"],
        "actors_annotated": query_one(
            """
            SELECT COUNT(DISTINCT actors.id) AS count
            FROM actors
            JOIN segment_actors ON segment_actors.actor_id = actors.id
            WHERE actors.project_id = ?
            """,
            (project_id,),
        )["count"],
        "top_relation_type": relation["relation_type"] if relation else "None yet",
        "top_feature_type": feature["feature_type"] if feature else "None yet",
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
            COUNT(DISTINCT segments.id) AS segment_count,
            COUNT(DISTINCT coded_segments.id) AS coded_segment_count
        FROM documents
        LEFT JOIN segments ON segments.document_id = documents.id
        LEFT JOIN (
            SELECT DISTINCT segments.id, segments.document_id
            FROM segments
            JOIN segment_codes ON segment_codes.segment_id = segments.id
        ) AS coded_segments ON coded_segments.document_id = documents.id
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


def get_latest_open_codes(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            codes.id,
            codes.name,
            codes.color,
            codes.created_at,
            COUNT(segment_codes.segment_id) AS assigned_segment_count
        FROM codes
        LEFT JOIN segment_codes ON segment_codes.code_id = codes.id
        WHERE codes.project_id = ? AND codes.code_type = 'open'
        GROUP BY codes.id
        ORDER BY datetime(codes.created_at) DESC, codes.id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def get_latest_coded_segments(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            segments.id,
            COALESCE(segments.name, '') AS name,
            segments.selected_text,
            documents.title AS document_title,
            GROUP_CONCAT(codes.name, ', ') AS code_names,
            segments.created_at
        FROM segments
        JOIN documents ON documents.id = segments.document_id
        JOIN segment_codes ON segment_codes.segment_id = segments.id
        JOIN codes ON codes.id = segment_codes.code_id
        WHERE documents.project_id = ? AND codes.code_type = 'open'
        GROUP BY segments.id
        ORDER BY datetime(segments.created_at) DESC, segments.id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def get_latest_memos(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT title, memo_type, status, created_at
        FROM memos
        WHERE project_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        LIMIT 5
        """,
        (project_id,),
    )


def get_codes_missing_definitions(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT id, name
        FROM codes
        WHERE project_id = ? AND code_type = 'open'
          AND (definition IS NULL OR TRIM(definition) = '')
        ORDER BY name COLLATE NOCASE
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


def get_segments_for_document(document_id: int, project_id: int | None = None) -> list[dict]:
    segment_rows = query_all(
        """
        SELECT id, document_id, COALESCE(name, '') AS name, selected_text, start_offset, end_offset,
               note, created_at, updated_at
        FROM segments
        WHERE document_id = ?
        ORDER BY start_offset ASC, end_offset ASC, id ASC
        """,
        (document_id,),
    )
    segments = []
    for row in segment_rows:
        segment = dict(row)
        segment["codes"] = get_codes_for_segment(row["id"])
        segment["discourse_markers"] = get_discourse_markers_for_segment(row["id"])
        segment["actors"] = get_segment_actors_for_segment(row["id"])
        segment["features"] = get_discourse_features_for_segment(row["id"])
        segment["memos"] = get_memos_for_entity(project_id, "segment", row["id"])
        segments.append(segment)
    return segments


def get_codes_for_segment(segment_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT codes.id, codes.name, codes.color
        FROM codes
        JOIN segment_codes ON segment_codes.code_id = codes.id
        WHERE segment_codes.segment_id = ? AND codes.code_type = 'open'
        ORDER BY codes.name COLLATE NOCASE
        """,
        (segment_id,),
    )


def get_discourse_markers_for_segment(segment_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT discourse_markers.id, discourse_markers.name,
               discourse_markers.marker_type, discourse_markers.color,
               segment_discourse_markers.note,
               segment_discourse_markers.created_at AS assigned_at
        FROM discourse_markers
        JOIN segment_discourse_markers
            ON segment_discourse_markers.marker_id = discourse_markers.id
        WHERE segment_discourse_markers.segment_id = ?
        ORDER BY discourse_markers.name COLLATE NOCASE
        """,
        (segment_id,),
    )


def get_segment_actors_for_segment(segment_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT segment_actors.id, segment_actors.relation_type, segment_actors.note,
               segment_actors.created_at, actors.name, actors.actor_type
        FROM segment_actors
        JOIN actors ON actors.id = segment_actors.actor_id
        WHERE segment_actors.segment_id = ?
        ORDER BY actors.name COLLATE NOCASE, segment_actors.relation_type
        """,
        (segment_id,),
    )


def get_discourse_features_for_segment(segment_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT id, feature_type, value, interpretation, created_at, updated_at
        FROM discourse_features
        WHERE segment_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        """,
        (segment_id,),
    )


def get_discourse_markers_for_project(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT discourse_markers.id, discourse_markers.name,
               discourse_markers.marker_type, discourse_markers.description,
               discourse_markers.color, discourse_markers.created_at,
               discourse_markers.updated_at,
               COUNT(segment_discourse_markers.segment_id) AS assigned_segment_count
        FROM discourse_markers
        LEFT JOIN segment_discourse_markers
            ON segment_discourse_markers.marker_id = discourse_markers.id
        WHERE discourse_markers.project_id = ?
        GROUP BY discourse_markers.id
        ORDER BY discourse_markers.name COLLATE NOCASE
        """,
        (project_id,),
    )


def get_discourse_marker_for_project(
    marker_id: int, project_id: int
) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT id, project_id, name, marker_type, description, color,
               created_at, updated_at
        FROM discourse_markers
        WHERE id = ? AND project_id = ?
        """,
        (marker_id, project_id),
    )


def get_actors_for_project(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT actors.id, actors.name, actors.actor_type, actors.description,
               actors.created_at, actors.updated_at,
               COUNT(segment_actors.id) AS annotation_count
        FROM actors
        LEFT JOIN segment_actors ON segment_actors.actor_id = actors.id
        WHERE actors.project_id = ?
        GROUP BY actors.id
        ORDER BY actors.name COLLATE NOCASE
        """,
        (project_id,),
    )


def get_actor_for_project(actor_id: int, project_id: int) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT id, project_id, name, actor_type, description, created_at, updated_at
        FROM actors
        WHERE id = ? AND project_id = ?
        """,
        (actor_id, project_id),
    )


def segment_has_discourse_marker(segment_id: int, marker_id: int) -> bool:
    return (
        query_one(
            """
            SELECT 1
            FROM segment_discourse_markers
            WHERE segment_id = ? AND marker_id = ?
            """,
            (segment_id, marker_id),
        )
        is not None
    )


def get_segment_actor_for_project(
    segment_actor_id: int, segment_id: int, project_id: int
) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT segment_actors.id, segment_actors.segment_id, segment_actors.actor_id,
               segment_actors.relation_type, actors.name AS actor_name,
               segments.document_id, COALESCE(segments.name, '') AS segment_name
        FROM segment_actors
        JOIN actors ON actors.id = segment_actors.actor_id
        JOIN segments ON segments.id = segment_actors.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE segment_actors.id = ? AND segment_actors.segment_id = ?
          AND documents.project_id = ?
        """,
        (segment_actor_id, segment_id, project_id),
    )


def get_discourse_feature_for_project(
    feature_id: int, segment_id: int, project_id: int
) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT discourse_features.id, discourse_features.segment_id,
               discourse_features.feature_type, discourse_features.value,
               segments.document_id, COALESCE(segments.name, '') AS segment_name
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE discourse_features.id = ? AND discourse_features.segment_id = ?
          AND documents.project_id = ?
        """,
        (feature_id, segment_id, project_id),
    )


def get_cda_counts(project_id: int) -> dict[str, int]:
    return {
        "markers": query_one(
            "SELECT COUNT(*) AS count FROM discourse_markers WHERE project_id = ?",
            (project_id,),
        )["count"],
        "actors": query_one(
            "SELECT COUNT(*) AS count FROM actors WHERE project_id = ?", (project_id,)
        )["count"],
        "segments_with_markers": query_one(
            """
            SELECT COUNT(DISTINCT segments.id) AS count
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            JOIN segment_discourse_markers
                ON segment_discourse_markers.segment_id = segments.id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        )["count"],
        "segments_with_actor_annotations": query_one(
            """
            SELECT COUNT(DISTINCT segments.id) AS count
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            JOIN segment_actors ON segment_actors.segment_id = segments.id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        )["count"],
        "metaphor_features": get_feature_type_count(project_id, "metaphor"),
        "presupposition_features": get_feature_type_count(project_id, "presupposition"),
        "legitimation_features": get_feature_type_count(project_id, "legitimation"),
    }


def get_feature_type_count(project_id: int, feature_type: str) -> int:
    return query_one(
        """
        SELECT COUNT(*) AS count
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ? AND discourse_features.feature_type = ?
        """,
        (project_id, feature_type),
    )["count"]


def get_discourse_features_for_project(project_id: int, filters: dict) -> list[sqlite3.Row]:
    sql = """
        SELECT discourse_features.id, discourse_features.feature_type,
               discourse_features.value, discourse_features.interpretation,
               discourse_features.created_at,
               COALESCE(segments.name, '') AS segment_name,
               segments.selected_text, documents.id AS document_id,
               documents.title AS document_title
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
    """
    params = [project_id]
    if filters.get("feature_type") in DISCOURSE_FEATURE_TYPES:
        sql += " AND discourse_features.feature_type = ?"
        params.append(filters["feature_type"])
    if filters.get("document_id", "").isdigit():
        sql += " AND documents.id = ?"
        params.append(int(filters["document_id"]))
    sql += " ORDER BY datetime(discourse_features.created_at) DESC, discourse_features.id DESC"
    return query_all(sql, tuple(params))


def get_discourse_feature_counts(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT discourse_features.feature_type, COUNT(*) AS count
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        GROUP BY discourse_features.feature_type
        ORDER BY count DESC, discourse_features.feature_type
        """,
        (project_id,),
    )


def get_voice_silence_report(project_id: int) -> list[dict]:
    rows = query_all(
        """
        SELECT actors.id, actors.name, actors.actor_type,
               segment_actors.relation_type,
               COUNT(segment_actors.id) AS relation_count,
               COUNT(DISTINCT segment_actors.segment_id) AS segment_count,
               COUNT(DISTINCT documents.id) AS document_count
        FROM actors
        LEFT JOIN segment_actors ON segment_actors.actor_id = actors.id
        LEFT JOIN segments ON segments.id = segment_actors.segment_id
        LEFT JOIN documents ON documents.id = segments.document_id
        WHERE actors.project_id = ?
        GROUP BY actors.id, segment_actors.relation_type
        ORDER BY actors.name COLLATE NOCASE
        """,
        (project_id,),
    )
    actors: dict[int, dict] = {}
    for index, row in enumerate(rows):
        actor = actors.setdefault(
            row["id"],
            {
                "id": row["id"],
                "name": row["name"],
                "actor_type": row["actor_type"],
                "total_annotations": 0,
                "document_count": 0,
                "segment_count": 0,
                **{relation: 0 for relation in ACTOR_RELATION_TYPES},
            },
        )
        if row["relation_type"] in ACTOR_RELATION_TYPES:
            actor[row["relation_type"]] = row["relation_count"]
            actor["total_annotations"] += row["relation_count"]
            actor["document_count"] += row["document_count"]
            actor["segment_count"] += row["segment_count"]

    for actor in actors.values():
        totals = query_one(
            """
            SELECT COUNT(DISTINCT documents.id) AS document_count,
                   COUNT(DISTINCT segments.id) AS segment_count
            FROM actors
            LEFT JOIN segment_actors ON segment_actors.actor_id = actors.id
            LEFT JOIN segments ON segments.id = segment_actors.segment_id
            LEFT JOIN documents ON documents.id = segments.document_id
            WHERE actors.id = ? AND actors.project_id = ?
            """,
            (actor["id"], project_id),
        )
        actor["document_count"] = totals["document_count"]
        actor["segment_count"] = totals["segment_count"]
    return list(actors.values())


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


def get_open_codes_for_project(project_id: int) -> list[sqlite3.Row]:
    return get_codes_for_project(project_id, "open")


def get_codes_for_project(project_id: int, code_type_filter: str = "all") -> list[sqlite3.Row]:
    where = "codes.project_id = ?"
    params = [project_id]
    if code_type_filter in {"open", "axial", "category"}:
        where += " AND codes.code_type = ?"
        params.append(code_type_filter)
    return query_all(
        f"""
        SELECT
            codes.id,
            codes.name,
            codes.description,
            codes.code_type,
            codes.color,
            codes.parent_id,
            parent.name AS parent_name,
            codes.definition,
            codes.include_when,
            codes.exclude_when,
            codes.example,
            codes.analytical_note,
            codes.created_at,
            codes.updated_at,
            COUNT(segment_codes.segment_id) AS assigned_segment_count,
            CASE
                WHEN codes.definition IS NULL OR TRIM(codes.definition) = '' THEN 'Missing definition'
                WHEN TRIM(COALESCE(codes.include_when, '')) != ''
                 AND TRIM(COALESCE(codes.exclude_when, '')) != ''
                 AND TRIM(COALESCE(codes.example, '')) != '' THEN 'Detailed entry'
                ELSE 'Basic entry'
            END AS completeness
        FROM codes
        LEFT JOIN segment_codes ON segment_codes.code_id = codes.id
        LEFT JOIN codes AS parent ON parent.id = codes.parent_id
        WHERE {where}
        GROUP BY codes.id
        ORDER BY codes.code_type, codes.name COLLATE NOCASE
        """,
        tuple(params),
    )


def get_code_for_project(code_id: int, project_id: int) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT id, project_id, name, description, code_type, color, parent_id,
               definition, include_when, exclude_when, example, analytical_note,
               gt_conditions, gt_context, gt_actions_interactions, gt_consequences,
               gt_properties, gt_dimensions, gt_theoretical_note,
               created_at, updated_at
        FROM codes
        WHERE id = ? AND project_id = ?
        """,
        (code_id, project_id),
    )


def segment_has_code(segment_id: int, code_id: int) -> bool:
    return (
        query_one(
            """
            SELECT 1
            FROM segment_codes
            WHERE segment_id = ? AND code_id = ?
            """,
            (segment_id, code_id),
        )
        is not None
    )


def create_gt_code(code_type: str, action: str, detail_prefix: str):
    active_project = get_active_project()
    name = request.form.get("name", "").strip()
    description = request.form.get("description", "").strip()
    color = normalize_code_color(request.form.get("color", ""))
    if not name:
        flash("Code name is required.", "error")
        return redirect(url_for("gt_workspace"))
    code_id = execute_write(
        """
        INSERT INTO codes (project_id, name, description, code_type, color)
        VALUES (?, ?, ?, ?, ?)
        """,
        (active_project["id"], name, description, code_type, color),
    )
    log_action(active_project["id"], "code", code_id, action, f"{detail_prefix}: {name}")
    flash(f"{detail_prefix}: {name}", "success")
    return redirect(url_for("gt_workspace"))


def assign_code_parent(
    child_id: int,
    child_type: str,
    parent_type: str,
    form_field: str,
    action: str,
    success_template: str,
):
    active_project = get_active_project()
    child = get_code_for_project(child_id, active_project["id"])
    parent_id_raw = request.form.get(form_field, "").strip()
    if child is None or child["code_type"] != child_type or not parent_id_raw.isdigit():
        flash("Invalid GT hierarchy operation.", "error")
        return redirect(url_for("gt_workspace"))
    parent_id = int(parent_id_raw)
    valid, message = validate_parent_assignment(
        child_id, child_type, parent_id, active_project["id"]
    )
    if not valid:
        flash(message, "error")
        return redirect(url_for("gt_workspace"))
    parent = get_code_for_project(parent_id, active_project["id"])
    execute_write(
        "UPDATE codes SET parent_id = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (parent_id, child_id),
    )
    details = success_template.format(child=child["name"], parent=parent["name"])
    log_action(active_project["id"], "code", child_id, action, details)
    flash(details, "success")
    return redirect(url_for("gt_workspace"))


def unassign_code_parent(code_id: int, expected_type: str, action: str):
    active_project = get_active_project()
    code = get_code_for_project(code_id, active_project["id"])
    if code is None or code["code_type"] != expected_type:
        flash("Invalid GT hierarchy operation.", "error")
        return redirect(url_for("gt_workspace"))
    execute_write(
        "UPDATE codes SET parent_id = NULL, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (code_id,),
    )
    log_action(
        active_project["id"],
        "code",
        code_id,
        action,
        f"Unassigned {expected_type} code: {code['name']}",
    )
    flash(f"Unassigned code: {code['name']}", "success")
    return redirect(url_for("gt_workspace"))


def validate_parent_assignment(
    code_id: int, child_type: str, parent_id: int, project_id: int
) -> tuple[bool, str]:
    if code_id == parent_id:
        return False, "Invalid GT hierarchy operation."
    parent = get_code_for_project(parent_id, project_id)
    expected_parent = {"open": "axial", "axial": "category"}.get(child_type)
    if expected_parent is None or parent is None or parent["code_type"] != expected_parent:
        return False, "Invalid GT hierarchy operation."
    return True, ""


def get_parent_code(parent_id: int, project_id: int) -> sqlite3.Row | None:
    return get_code_for_project(parent_id, project_id)


def get_child_codes(parent_id: int, project_id: int, code_type: str) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT id, name, description, color, code_type
        FROM codes
        WHERE parent_id = ? AND project_id = ? AND code_type = ?
        ORDER BY name COLLATE NOCASE
        """,
        (parent_id, project_id, code_type),
    )


def get_open_codes_under_category(category_id: int, project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT open_code.id, open_code.name, open_code.description, open_code.color,
               axial_code.name AS axial_name
        FROM codes AS axial_code
        JOIN codes AS open_code ON open_code.parent_id = axial_code.id
        WHERE axial_code.parent_id = ? AND axial_code.project_id = ?
          AND open_code.project_id = ? AND axial_code.code_type = 'axial'
          AND open_code.code_type = 'open'
        ORDER BY axial_code.name COLLATE NOCASE, open_code.name COLLATE NOCASE
        """,
        (category_id, project_id, project_id),
    )


def get_gt_open_codes(project_id: int) -> list[dict]:
    rows = query_all(
        """
        SELECT codes.*, parent.name AS parent_name,
               COUNT(segment_codes.segment_id) AS assigned_segment_count
        FROM codes
        LEFT JOIN codes AS parent ON parent.id = codes.parent_id
        LEFT JOIN segment_codes ON segment_codes.code_id = codes.id
        WHERE codes.project_id = ? AND codes.code_type = 'open'
        GROUP BY codes.id
        ORDER BY codes.name COLLATE NOCASE
        """,
        (project_id,),
    )
    return [dict(row) for row in rows]


def get_gt_axial_codes(project_id: int) -> list[dict]:
    rows = query_all(
        """
        SELECT codes.*, parent.name AS parent_name
        FROM codes
        LEFT JOIN codes AS parent ON parent.id = codes.parent_id
        WHERE codes.project_id = ? AND codes.code_type = 'axial'
        ORDER BY codes.name COLLATE NOCASE
        """,
        (project_id,),
    )
    result = []
    for index, row in enumerate(rows):
        code = dict(row)
        code["child_open_count"] = len(get_child_codes(row["id"], project_id, "open"))
        code["hierarchy_segment_count"] = len(get_hierarchy_segments_for_code(row, project_id))
        result.append(code)
    return result


def get_gt_categories(project_id: int) -> list[dict]:
    rows = query_all(
        """
        SELECT *
        FROM codes
        WHERE project_id = ? AND code_type = 'category'
        ORDER BY name COLLATE NOCASE
        """,
        (project_id,),
    )
    result = []
    for row in rows:
        code = dict(row)
        axial_children = get_child_codes(row["id"], project_id, "axial")
        code["child_axial_count"] = len(axial_children)
        code["child_open_count"] = sum(
            len(get_child_codes(axial["id"], project_id, "open"))
            for axial in axial_children
        )
        code["hierarchy_segment_count"] = len(get_hierarchy_segments_for_code(row, project_id))
        result.append(code)
    return result


def get_code_usage_count(code_id: int) -> int:
    return query_one(
        "SELECT COUNT(*) AS count FROM segment_codes WHERE code_id = ?",
        (code_id,),
    )["count"]


def get_segments_for_code(code_id: int, project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            segments.id,
            COALESCE(segments.name, '') AS name,
            segments.selected_text,
            segments.note,
            documents.id AS document_id,
            documents.title AS document_title
        FROM segments
        JOIN documents ON documents.id = segments.document_id
        JOIN segment_codes ON segment_codes.segment_id = segments.id
        WHERE segment_codes.code_id = ? AND documents.project_id = ?
        ORDER BY documents.title COLLATE NOCASE, segments.start_offset
        """,
        (code_id, project_id),
    )


def get_hierarchy_segments_for_code(code: sqlite3.Row | dict, project_id: int) -> list[sqlite3.Row]:
    if code["code_type"] == "open":
        return get_segments_for_code(code["id"], project_id)
    if code["code_type"] == "axial":
        return query_all(
            """
            SELECT DISTINCT segments.id, COALESCE(segments.name, '') AS name,
                   segments.selected_text, segments.note, documents.id AS document_id,
                   documents.title AS document_title
            FROM codes AS open_code
            JOIN segment_codes ON segment_codes.code_id = open_code.id
            JOIN segments ON segments.id = segment_codes.segment_id
            JOIN documents ON documents.id = segments.document_id
            WHERE open_code.parent_id = ? AND documents.project_id = ?
            ORDER BY documents.title COLLATE NOCASE, segments.start_offset
            """,
            (code["id"], project_id),
        )
    if code["code_type"] == "category":
        return query_all(
            """
            SELECT DISTINCT segments.id, COALESCE(segments.name, '') AS name,
                   segments.selected_text, segments.note, documents.id AS document_id,
                   documents.title AS document_title
            FROM codes AS axial_code
            JOIN codes AS open_code ON open_code.parent_id = axial_code.id
            JOIN segment_codes ON segment_codes.code_id = open_code.id
            JOIN segments ON segments.id = segment_codes.segment_id
            JOIN documents ON documents.id = segments.document_id
            WHERE axial_code.parent_id = ? AND documents.project_id = ?
            ORDER BY documents.title COLLATE NOCASE, segments.start_offset
            """,
            (code["id"], project_id),
        )
    return []


def get_compare_code(code_id_raw: str | None, project_id: int) -> dict | None:
    if not code_id_raw or not code_id_raw.isdigit():
        return None
    code = get_code_for_project(int(code_id_raw), project_id)
    if code is None or code["code_type"] != "open":
        return None
    result = dict(code)
    result["segments"] = get_segments_for_code(code["id"], project_id)
    result["usage_count"] = len(result["segments"])
    return result


def get_gt_structure_preview(project_id: int) -> dict[str, int]:
    return {
        "open_without_axial": query_one(
            """
            SELECT COUNT(*) AS count FROM codes
            WHERE project_id = ? AND code_type = 'open' AND parent_id IS NULL
            """,
            (project_id,),
        )["count"],
        "axial_without_category": query_one(
            """
            SELECT COUNT(*) AS count FROM codes
            WHERE project_id = ? AND code_type = 'axial' AND parent_id IS NULL
            """,
            (project_id,),
        )["count"],
        "categories": query_one(
            """
            SELECT COUNT(*) AS count FROM codes
            WHERE project_id = ? AND code_type = 'category'
            """,
            (project_id,),
        )["count"],
    }


def get_documents_for_memo_links(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT id, title
        FROM documents
        WHERE project_id = ?
        ORDER BY title COLLATE NOCASE
        """,
        (project_id,),
    )


def get_segments_for_memo_links(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            segments.id,
            COALESCE(segments.name, '') AS name,
            segments.selected_text,
            documents.title AS document_title
        FROM segments
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        ORDER BY documents.title COLLATE NOCASE, segments.start_offset
        """,
        (project_id,),
    )


def get_memos_for_project(project_id: int, filters: dict) -> list[dict]:
    sql = """
        SELECT id, title, body, memo_type, linked_entity_type, linked_entity_id,
               status, created_at, updated_at
        FROM memos
        WHERE project_id = ?
    """
    params = [project_id]
    if filters.get("memo_type") in MEMO_TYPES:
        sql += " AND memo_type = ?"
        params.append(filters["memo_type"])
    if filters.get("status") in MEMO_STATUSES:
        sql += " AND status = ?"
        params.append(filters["status"])
    if filters.get("linked_entity_type") in {"project", "document", "segment", "code"}:
        sql += " AND linked_entity_type = ?"
        params.append(filters["linked_entity_type"])
    if filters.get("linked_entity_id", "").isdigit():
        sql += " AND linked_entity_id = ?"
        params.append(int(filters["linked_entity_id"]))
    sql += " ORDER BY datetime(created_at) DESC, id DESC"

    memos = []
    for row in query_all(sql, tuple(params)):
        memo = dict(row)
        memo["linked_entity_label"] = get_linked_entity_label(memo, project_id)
        memos.append(memo)
    return memos


def get_memo_for_project(memo_id: int, project_id: int) -> sqlite3.Row | None:
    return query_one(
        """
        SELECT id, project_id, title, body, memo_type, linked_entity_type,
               linked_entity_id, status, created_at, updated_at
        FROM memos
        WHERE id = ? AND project_id = ?
        """,
        (memo_id, project_id),
    )


def get_memos_for_entity(
    project_id: int | None, linked_entity_type: str, linked_entity_id: int
) -> list[sqlite3.Row]:
    if project_id is None:
        return query_all(
            """
            SELECT id, title, body, memo_type, status, created_at
            FROM memos
            WHERE linked_entity_type = ? AND linked_entity_id = ?
            ORDER BY datetime(created_at) DESC, id DESC
            """,
            (linked_entity_type, linked_entity_id),
        )
    return query_all(
        """
        SELECT id, title, body, memo_type, status, created_at
        FROM memos
        WHERE project_id = ? AND linked_entity_type = ? AND linked_entity_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        """,
        (project_id, linked_entity_type, linked_entity_id),
    )


def validate_memo_form(project_id: int) -> tuple[dict, str | None]:
    title = request.form.get("title", "").strip()
    body = request.form.get("body", "").strip()
    memo_type = request.form.get("memo_type", "").strip()
    status = request.form.get("status", "draft").strip()
    linked_target = request.form.get("linked_target", "").strip()
    linked_entity_type = request.form.get("linked_entity_type", "").strip() or None
    linked_entity_id_raw = request.form.get("linked_entity_id", "").strip()
    if linked_target:
        try:
            linked_entity_type, linked_entity_id_raw = linked_target.split(":", 1)
        except ValueError:
            return {}, "Invalid memo link."

    if not title or not body:
        return {}, "Memo title and body are required."
    if memo_type not in MEMO_TYPES:
        return {}, "Invalid memo type."
    if status not in MEMO_STATUSES:
        return {}, "Invalid memo status."

    linked_entity_id = None
    if linked_entity_type or linked_entity_id_raw:
        if linked_entity_type not in {"project", "document", "segment", "code"}:
            return {}, "Invalid memo link."
        if not linked_entity_id_raw.isdigit():
            return {}, "Invalid memo link."
        linked_entity_id = int(linked_entity_id_raw)
        if not linked_entity_exists(project_id, linked_entity_type, linked_entity_id):
            return {}, "Invalid memo link."

    return {
        "title": title,
        "body": body,
        "memo_type": memo_type,
        "status": status,
        "linked_entity_type": linked_entity_type,
        "linked_entity_id": linked_entity_id,
    }, None


def linked_entity_exists(project_id: int, entity_type: str, entity_id: int) -> bool:
    if entity_type == "project":
        return entity_id == project_id
    if entity_type == "document":
        return get_document_for_project(entity_id, project_id) is not None
    if entity_type == "segment":
        return get_segment_for_project(entity_id, project_id) is not None
    if entity_type == "code":
        return get_code_for_project(entity_id, project_id) is not None
    return False


def get_linked_entity_label(memo: dict, project_id: int) -> str:
    entity_type = memo.get("linked_entity_type")
    entity_id = memo.get("linked_entity_id")
    if not entity_type or entity_id is None:
        return "None"
    if entity_type == "project":
        project = get_active_project()
        return f"Project: {project['name']}"
    if entity_type == "document":
        document = get_document_for_project(entity_id, project_id)
        return f"Document: {document['title']}" if document else "Document: missing"
    if entity_type == "segment":
        segment = get_segment_for_project(entity_id, project_id)
        if segment is None:
            return "Segment: missing"
        label = segment["name"] or truncate_text(segment["selected_text"], 48)
        return f"Segment: {label}"
    if entity_type == "code":
        code = get_code_for_project(entity_id, project_id)
        return f"Code: {code['name']}" if code else "Code: missing"
    return "None"


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
    document_text: str, segments: list[dict]
) -> Markup:
    pieces = []
    cursor = 0

    for segment in segments:
        start = segment["start_offset"]
        end = segment["end_offset"]
        if start < cursor or start < 0 or end > len(document_text) or end <= start:
            continue

        if segment["codes"]:
            color = segment["codes"][0]["color"]
        elif segment["discourse_markers"]:
            color = segment["discourse_markers"][0]["color"]
        else:
            color = "#fff0a8"
        style = f"--segment-color: {normalize_code_color(color)};"
        pieces.append(escape(document_text[cursor:start]))
        pieces.append(
            Markup(
                '<mark class="segment-highlight" data-segment-id="{}" style="{}">{}</mark>'
            ).format(
                segment["id"], style, escape(document_text[start:end])
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
        db.execute(
            f"DELETE FROM segment_discourse_markers WHERE segment_id IN ({placeholders})",
            tuple(segment_ids),
        )
        db.execute(
            f"DELETE FROM segment_actors WHERE segment_id IN ({placeholders})",
            tuple(segment_ids),
        )
        db.execute(
            f"DELETE FROM discourse_features WHERE segment_id IN ({placeholders})",
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


def redirect_after_code_change(document_id: str):
    if document_id.isdigit():
        active_project = get_active_project()
        document = get_document_for_project(int(document_id), active_project["id"])
        if document is not None:
            return redirect(url_for("document_view", document_id=document["id"]))
    return redirect(url_for("codes"))


def safe_next_url(next_url: str) -> str | None:
    if next_url.startswith("/") and not next_url.startswith("//"):
        return next_url
    return None


def truncate_text(text: str, length: int = 80) -> str:
    text = " ".join((text or "").split())
    if len(text) <= length:
        return text
    return text[: length - 3] + "..."


def get_dashboard_export_links() -> list[dict[str, str]]:
    return [
        {"label": "Codebook", "endpoint": "export_codebook_markdown"},
        {"label": "Coded segments", "endpoint": "export_coded_segments_csv"},
        {"label": "Memos", "endpoint": "export_memos_markdown"},
        {"label": "Project package", "endpoint": "export_project_package"},
    ]


def export_timestamp() -> str:
    return datetime.now().isoformat(timespec="seconds")


def empty(value) -> str:
    return "" if value is None else str(value)


def row_to_dict(row: sqlite3.Row) -> dict:
    return {key: row[key] for key in row.keys()}


def rows_to_dicts(rows: list[sqlite3.Row]) -> list[dict]:
    return [row_to_dict(row) for row in rows]


def download_text(content: str, filename: str, mimetype: str) -> Response:
    return Response(
        content,
        content_type=mimetype,
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def download_binary(content: bytes, filename: str, mimetype: str) -> Response:
    return Response(
        content,
        mimetype=mimetype,
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def make_csv(headers: list[str], rows: list[dict]) -> str:
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore")
    writer.writeheader()
    for row in rows:
        writer.writerow({header: empty(row.get(header)) for header in headers})
    return output.getvalue()


def markdown_blockquote(text: str) -> str:
    text = empty(text).strip()
    if not text:
        return "> "
    return "\n".join(f"> {line}" if line else ">" for line in text.splitlines())


def build_codebook_markdown(active_project: sqlite3.Row) -> str:
    project_id = active_project["id"]
    lines = [
        "# discourseLab Codebook",
        "",
        f"Project: {active_project['name']}",
        f"Exported: {export_timestamp()}",
        "",
    ]
    codes = query_all(
        """
        SELECT
            codes.*,
            parent.name AS parent_name,
            COUNT(DISTINCT segment_codes.segment_id) AS usage_count,
            COUNT(DISTINCT memos.id) AS linked_memo_count
        FROM codes
        LEFT JOIN segment_codes ON segment_codes.code_id = codes.id
        LEFT JOIN codes AS parent ON parent.id = codes.parent_id
        LEFT JOIN memos ON memos.project_id = codes.project_id
            AND memos.linked_entity_type = 'code'
            AND memos.linked_entity_id = codes.id
        WHERE codes.project_id = ?
        GROUP BY codes.id
        ORDER BY codes.code_type COLLATE NOCASE, codes.name COLLATE NOCASE
        """,
        (project_id,),
    )
    if not codes:
        lines.append("No codes created yet.")
        lines.append("")
        return "\n".join(lines)

    for code in codes:
        lines.extend(
            [
                f"## {code['name']}",
                "",
                f"- Type: {code['code_type']}",
                f"- Parent: {code['parent_name'] or 'None'}",
                f"- Color: {code['color'] or ''}",
                f"- Usage count: {code['usage_count']}",
                f"- Linked memo count: {code['linked_memo_count']}",
                "",
                f"**Description:** {code['description'] or ''}",
                "",
                f"**Definition:** {code['definition'] or ''}",
                "",
                f"**Include when:** {code['include_when'] or ''}",
                "",
                f"**Exclude when:** {code['exclude_when'] or ''}",
                "",
                f"**Example:** {code['example'] or ''}",
                "",
                f"**Analytical note:** {code['analytical_note'] or ''}",
                "",
            ]
        )
        if code["code_type"] in {"axial", "category"}:
            lines.extend(
                [
                    f"**GT conditions:** {code['gt_conditions'] or ''}",
                    "",
                    f"**GT context:** {code['gt_context'] or ''}",
                    "",
                    f"**GT actions/interactions:** {code['gt_actions_interactions'] or ''}",
                    "",
                    f"**GT consequences:** {code['gt_consequences'] or ''}",
                    "",
                    f"**GT properties:** {code['gt_properties'] or ''}",
                    "",
                    f"**GT dimensions:** {code['gt_dimensions'] or ''}",
                    "",
                    f"**GT theoretical note:** {code['gt_theoretical_note'] or ''}",
                    "",
                ]
            )
    return "\n".join(lines)


def get_coded_segment_export_rows(project_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            documents.id AS document_id,
            documents.title AS document_title,
            documents.original_filename AS document_original_filename,
            segments.id AS segment_id,
            COALESCE(segments.name, '') AS segment_title,
            segments.selected_text,
            segments.note AS segment_note,
            segments.start_offset,
            segments.end_offset,
            codes.id AS code_id,
            codes.name AS code_name,
            codes.code_type,
            codes.parent_id AS code_parent_id,
            CASE
                WHEN codes.code_type = 'open' THEN axial.name
                WHEN codes.code_type = 'axial' THEN codes.name
                ELSE NULL
            END AS axial_code_name,
            CASE
                WHEN codes.code_type = 'open' THEN category.name
                WHEN codes.code_type = 'category' THEN codes.name
                ELSE NULL
            END AS category_name,
            codes.description AS code_description,
            codes.definition AS code_definition,
            segments.created_at
        FROM segment_codes
        JOIN segments ON segments.id = segment_codes.segment_id
        JOIN documents ON documents.id = segments.document_id
        JOIN codes ON codes.id = segment_codes.code_id
        LEFT JOIN codes AS axial ON axial.id = codes.parent_id
            AND axial.code_type = 'axial'
        LEFT JOIN codes AS category ON category.id = axial.parent_id
            AND category.code_type = 'category'
        WHERE documents.project_id = ? AND codes.project_id = ?
        ORDER BY documents.title COLLATE NOCASE, segments.start_offset, codes.name COLLATE NOCASE
        """,
        (project_id, project_id),
    )


def generate_coded_segments_csv(active_project: sqlite3.Row) -> str:
    headers = [
        "project_name",
        "document_id",
        "document_title",
        "document_original_filename",
        "segment_id",
        "segment_title",
        "selected_text",
        "segment_note",
        "start_offset",
        "end_offset",
        "code_id",
        "code_name",
        "code_type",
        "code_parent_id",
        "axial_code_name",
        "category_name",
        "code_description",
        "code_definition",
        "created_at",
    ]
    rows = []
    for row in get_coded_segment_export_rows(active_project["id"]):
        data = row_to_dict(row)
        data["project_name"] = active_project["name"]
        rows.append(data)
    return make_csv(headers, rows)


def generate_coded_segments_markdown(active_project: sqlite3.Row) -> str:
    rows = get_coded_segment_export_rows(active_project["id"])
    lines = [
        "# Coded Segments",
        "",
        f"Project: {active_project['name']}",
        f"Exported: {export_timestamp()}",
        "",
    ]
    if not rows:
        lines.extend(["No coded segments found.", ""])
        return "\n".join(lines)

    current_document = None
    current_segment = None
    for index, row in enumerate(rows):
        if row["document_id"] != current_document:
            current_document = row["document_id"]
            current_segment = None
            lines.extend([f"## Document: {row['document_title']}", ""])
        if row["segment_id"] != current_segment:
            current_segment = row["segment_id"]
            segment_label = row["segment_title"] or f"Segment {row['segment_id']}"
            lines.extend(
                [
                    f"### Segment: {segment_label}",
                    "",
                    "Selected text:",
                    "",
                    markdown_blockquote(row["selected_text"]),
                    "",
                    "Segment note:",
                    "",
                    row["segment_note"] or "",
                    "",
                    "Codes:",
                ]
            )
        lines.append(f"- {row['code_name']}")
        lines.append(f"  - Type: {row['code_type']}")
        if row["axial_code_name"]:
            lines.append(f"  - Axial parent: {row['axial_code_name']}")
        if row["category_name"]:
            lines.append(f"  - Category parent: {row['category_name']}")

        next_is_new_segment = True
        next_index = index + 1
        if next_index < len(rows):
            next_is_new_segment = rows[next_index]["segment_id"] != current_segment
        if next_is_new_segment:
            memos = get_memos_for_entity(active_project["id"], "segment", row["segment_id"])
            lines.extend(["", "Memos linked to this segment:"])
            if memos:
                for memo in memos:
                    lines.append(
                        f"- {memo['title']}, {memo['status']}, {memo['memo_type']}"
                    )
                    lines.append(f"  - {empty(memo['body']).replace(chr(10), ' ')}")
            else:
                lines.append("- None")
            lines.append("")
    return "\n".join(lines)


def generate_memos_markdown(active_project: sqlite3.Row) -> str:
    memos = get_memos_for_project(
        active_project["id"],
        {"memo_type": "", "status": "", "linked_entity_type": "", "linked_entity_id": ""},
    )
    lines = [
        "# Memos",
        "",
        f"Project: {active_project['name']}",
        f"Exported: {export_timestamp()}",
        "",
    ]
    if not memos:
        lines.extend(["No memos found.", ""])
        return "\n".join(lines)
    current_type = None
    current_status = None
    for memo in sorted(memos, key=lambda item: (item["memo_type"], item["status"], item["title"].lower())):
        if memo["memo_type"] != current_type:
            current_type = memo["memo_type"]
            current_status = None
            lines.extend([f"## {MEMO_TYPES.get(current_type, current_type)}", ""])
        if memo["status"] != current_status:
            current_status = memo["status"]
            lines.extend([f"### {MEMO_STATUSES.get(current_status, current_status)}", ""])
        lines.extend(
            [
                f"#### {memo['title']}",
                "",
                f"Status: {MEMO_STATUSES.get(memo['status'], memo['status'])}",
                f"Linked to: {memo['linked_entity_label']}",
                f"Created: {memo['created_at']}",
                f"Updated: {memo['updated_at']}",
                "",
                memo["body"] or "",
                "",
            ]
        )
    return "\n".join(lines)


def generate_cda_features_csv(active_project: sqlite3.Row) -> str:
    headers = [
        "project_name",
        "document_id",
        "document_title",
        "segment_id",
        "segment_title",
        "selected_text",
        "feature_id",
        "feature_type",
        "value",
        "interpretation",
        "created_at",
        "updated_at",
    ]
    rows = []
    feature_rows = query_all(
        """
        SELECT documents.id AS document_id, documents.title AS document_title,
               segments.id AS segment_id, COALESCE(segments.name, '') AS segment_title,
               segments.selected_text, discourse_features.id AS feature_id,
               discourse_features.feature_type, discourse_features.value,
               discourse_features.interpretation, discourse_features.created_at,
               discourse_features.updated_at
        FROM discourse_features
        JOIN segments ON segments.id = discourse_features.segment_id
        JOIN documents ON documents.id = segments.document_id
        WHERE documents.project_id = ?
        ORDER BY documents.title COLLATE NOCASE, segments.start_offset,
                 discourse_features.feature_type
        """,
        (active_project["id"],),
    )
    for row in feature_rows:
        data = row_to_dict(row)
        data["project_name"] = active_project["name"]
        rows.append(data)
    return make_csv(headers, rows)


def generate_voice_silence_csv(active_project: sqlite3.Row) -> str:
    headers = [
        "project_name",
        "actor_id",
        "actor_name",
        "actor_type",
        "actor_description",
        *ACTOR_RELATION_TYPES.keys(),
        "total_annotations",
        "segment_count",
        "document_count",
    ]
    actors_by_id = {
        actor["id"]: actor
        for actor in query_all(
            """
            SELECT id, name, actor_type, description
            FROM actors
            WHERE project_id = ?
            ORDER BY name COLLATE NOCASE
            """,
            (active_project["id"],),
        )
    }
    rows = []
    for report_row in get_voice_silence_report(active_project["id"]):
        actor = actors_by_id.get(report_row["id"])
        row = {
            "project_name": active_project["name"],
            "actor_id": report_row["id"],
            "actor_name": report_row["name"],
            "actor_type": report_row["actor_type"],
            "actor_description": actor["description"] if actor else "",
            "total_annotations": report_row["total_annotations"],
            "segment_count": report_row["segment_count"],
            "document_count": report_row["document_count"],
        }
        for relation in ACTOR_RELATION_TYPES:
            row[relation] = report_row[relation]
        rows.append(row)
    return make_csv(headers, rows)


def generate_gt_hierarchy_markdown(active_project: sqlite3.Row) -> str:
    project_id = active_project["id"]
    categories = get_gt_categories(project_id)
    unassigned_axial = [
        code for code in get_gt_axial_codes(project_id) if code.get("parent_id") is None
    ]
    unassigned_open = [
        code for code in get_gt_open_codes(project_id) if code.get("parent_id") is None
    ]
    lines = [
        "# Grounded Theory Hierarchy",
        "",
        f"Project: {active_project['name']}",
        f"Exported: {export_timestamp()}",
        "",
        "## Categories",
        "",
    ]
    if not categories:
        lines.extend(["No categories found.", ""])
    for category in categories:
        axial_children = get_child_codes(category["id"], project_id, "axial")
        lines.extend(
            [
                f"### Category: {category['name']}",
                "",
                f"Description: {category.get('description') or ''}",
                f"Definition: {category.get('definition') or ''}",
                f"Analytical note: {category.get('analytical_note') or ''}",
                f"GT theoretical note: {category.get('gt_theoretical_note') or ''}",
                "",
                "Child axial codes:",
                "",
            ]
        )
        if not axial_children:
            lines.extend(["None.", ""])
        for axial in axial_children:
            axial_full = get_code_for_project(axial["id"], project_id)
            open_children = get_child_codes(axial["id"], project_id, "open")
            lines.extend(
                [
                    f"#### Axial code: {axial_full['name']}",
                    "",
                    f"Description: {axial_full['description'] or ''}",
                    f"Conditions: {axial_full['gt_conditions'] or ''}",
                    f"Context: {axial_full['gt_context'] or ''}",
                    f"Actions/interactions: {axial_full['gt_actions_interactions'] or ''}",
                    f"Consequences: {axial_full['gt_consequences'] or ''}",
                    f"Properties: {axial_full['gt_properties'] or ''}",
                    f"Dimensions: {axial_full['gt_dimensions'] or ''}",
                    "",
                    "Child open codes:",
                ]
            )
            if open_children:
                for open_code in open_children:
                    usage_count = get_code_usage_count(open_code["id"])
                    lines.append(f"- {open_code['name']}")
                    lines.append(f"  - Description: {open_code['description'] or ''}")
                    lines.append(f"  - Assigned segment count: {usage_count}")
            else:
                lines.append("- None")
            lines.extend(["", "Representative segments:"])
            segments = get_hierarchy_segments_for_code(axial_full, project_id)[:5]
            if segments:
                for segment in segments:
                    lines.extend([markdown_blockquote(segment["selected_text"]), ""])
            else:
                lines.extend(["> None", ""])

    lines.extend(["## Unassigned axial codes", ""])
    if unassigned_axial:
        for code in unassigned_axial:
            lines.append(f"- {code['name']}")
    else:
        lines.append("None.")
    lines.extend(["", "## Unassigned open codes", ""])
    if unassigned_open:
        for code in unassigned_open:
            lines.append(f"- {code['name']} ({code['assigned_segment_count']} assigned segments)")
    else:
        lines.append("None.")
    lines.append("")
    return "\n".join(lines)


def generate_project_summary_markdown(active_project: sqlite3.Row) -> str:
    project_id = active_project["id"]
    counts = get_dashboard_counts(project_id)
    documents = get_documents_for_project(project_id)
    research_questions = query_all(
        """
        SELECT question, note, created_at
        FROM research_questions
        WHERE project_id = ?
        ORDER BY datetime(created_at) DESC, id DESC
        """,
        (project_id,),
    )
    top_codes = query_all(
        """
        SELECT codes.name, codes.code_type, COUNT(segment_codes.segment_id) AS count
        FROM codes
        LEFT JOIN segment_codes ON segment_codes.code_id = codes.id
        WHERE codes.project_id = ?
        GROUP BY codes.id
        ORDER BY count DESC, codes.name COLLATE NOCASE
        LIMIT 10
        """,
        (project_id,),
    )
    top_actors = query_all(
        """
        SELECT actors.name, actors.actor_type, COUNT(segment_actors.id) AS count
        FROM actors
        LEFT JOIN segment_actors ON segment_actors.actor_id = actors.id
        WHERE actors.project_id = ?
        GROUP BY actors.id
        ORDER BY count DESC, actors.name COLLATE NOCASE
        LIMIT 10
        """,
        (project_id,),
    )
    top_features = get_discourse_feature_counts(project_id)[:10]
    important_memos = query_all(
        """
        SELECT title, memo_type, status, created_at
        FROM memos
        WHERE project_id = ? AND status IN ('important', 'use_in_article')
        ORDER BY datetime(created_at) DESC, id DESC
        LIMIT 10
        """,
        (project_id,),
    )
    audit_summary = query_all(
        """
        SELECT action, COUNT(*) AS count, MAX(created_at) AS latest_at
        FROM audit_log
        WHERE project_id = ?
        GROUP BY action
        ORDER BY count DESC, action
        """,
        (project_id,),
    )
    lines = [
        "# Project Summary",
        "",
        f"Project: {active_project['name']}",
        f"Description: {active_project['description'] or ''}",
        f"Exported: {export_timestamp()}",
        "",
        "## Counts",
        "",
    ]
    for key in [
        "documents",
        "segments",
        "open_codes",
        "axial_codes",
        "categories",
        "coded_segments",
        "memos",
        "cda_markers",
        "actors",
        "discourse_features",
        "relations",
    ]:
        lines.append(f"- {key.replace('_', ' ').title()}: {counts.get(key, 0)}")
    lines.extend(["", "## Documents", ""])
    if documents:
        for document in documents:
            lines.append(
                f"- {document['title']} ({document['file_type']}, {document['segment_count']} segments)"
            )
    else:
        lines.append("No documents imported.")
    lines.extend(["", "## Research Questions", ""])
    if research_questions:
        for question in research_questions:
            lines.append(f"- {question['question']}")
            if question["note"]:
                lines.append(f"  - {question['note']}")
    else:
        lines.append("No research questions recorded.")
    lines.extend(["", "## Top Codes by Segment Count", ""])
    lines.extend([f"- {row['name']} ({row['code_type']}): {row['count']}" for row in top_codes] or ["None."])
    lines.extend(["", "## Top Actors by Annotation Count", ""])
    lines.extend([f"- {row['name']} ({row['actor_type']}): {row['count']}" for row in top_actors] or ["None."])
    lines.extend(["", "## Top Discourse Feature Types", ""])
    lines.extend([f"- {row['feature_type']}: {row['count']}" for row in top_features] or ["None."])
    lines.extend(["", "## Latest Important Memos", ""])
    lines.extend([f"- {row['title']} ({row['memo_type']}, {row['status']}, {row['created_at']})" for row in important_memos] or ["None."])
    lines.extend(["", "## Audit Log Summary", ""])
    lines.extend([f"- {row['action']}: {row['count']} actions, latest {row['latest_at']}" for row in audit_summary] or ["No audit log entries."])
    lines.append("")
    return "\n".join(lines)


def generate_project_json(active_project: sqlite3.Row) -> str:
    project_id = active_project["id"]
    segment_ids = [
        row["id"]
        for row in query_all(
            """
            SELECT segments.id
            FROM segments
            JOIN documents ON documents.id = segments.document_id
            WHERE documents.project_id = ?
            """,
            (project_id,),
        )
    ]
    document_ids = [
        row["id"]
        for row in query_all("SELECT id FROM documents WHERE project_id = ?", (project_id,))
    ]
    code_ids = [
        row["id"] for row in query_all("SELECT id FROM codes WHERE project_id = ?", (project_id,))
    ]
    actor_ids = [
        row["id"] for row in query_all("SELECT id FROM actors WHERE project_id = ?", (project_id,))
    ]
    marker_ids = [
        row["id"]
        for row in query_all("SELECT id FROM discourse_markers WHERE project_id = ?", (project_id,))
    ]

    def rows_for_ids(sql: str, ids: list[int]) -> list[dict]:
        if not ids:
            return []
        placeholders = ",".join("?" for _ in ids)
        return rows_to_dicts(query_all(sql.format(placeholders=placeholders), tuple(ids)))

    data = {
        "project": row_to_dict(active_project),
        "documents": rows_to_dicts(query_all("SELECT * FROM documents WHERE project_id = ? ORDER BY id", (project_id,))),
        "tags": rows_to_dicts(query_all("SELECT * FROM tags WHERE project_id = ? ORDER BY id", (project_id,))),
        "document_tags": rows_for_ids("SELECT * FROM document_tags WHERE document_id IN ({placeholders}) ORDER BY document_id, tag_id", document_ids),
        "codes": rows_to_dicts(query_all("SELECT * FROM codes WHERE project_id = ? ORDER BY id", (project_id,))),
        "segments": rows_for_ids("SELECT * FROM segments WHERE id IN ({placeholders}) ORDER BY id", segment_ids),
        "segment_codes": rows_for_ids("SELECT * FROM segment_codes WHERE segment_id IN ({placeholders}) ORDER BY segment_id, code_id", segment_ids),
        "memos": rows_to_dicts(query_all("SELECT * FROM memos WHERE project_id = ? ORDER BY id", (project_id,))),
        "relations": rows_to_dicts(query_all("SELECT * FROM relations WHERE project_id = ? ORDER BY id", (project_id,))),
        "research_questions": rows_to_dicts(query_all("SELECT * FROM research_questions WHERE project_id = ? ORDER BY id", (project_id,))),
        "discourse_markers": rows_to_dicts(query_all("SELECT * FROM discourse_markers WHERE project_id = ? ORDER BY id", (project_id,))),
        "segment_discourse_markers": rows_for_ids("SELECT * FROM segment_discourse_markers WHERE marker_id IN ({placeholders}) ORDER BY segment_id, marker_id", marker_ids),
        "actors": rows_to_dicts(query_all("SELECT * FROM actors WHERE project_id = ? ORDER BY id", (project_id,))),
        "segment_actors": rows_for_ids("SELECT * FROM segment_actors WHERE actor_id IN ({placeholders}) ORDER BY id", actor_ids),
        "discourse_features": rows_for_ids("SELECT * FROM discourse_features WHERE segment_id IN ({placeholders}) ORDER BY id", segment_ids),
        "audit_log": rows_to_dicts(query_all("SELECT * FROM audit_log WHERE project_id = ? ORDER BY id", (project_id,))),
    }
    return json.dumps(data, indent=2, ensure_ascii=False)


def generate_project_package_zip(active_project: sqlite3.Row) -> bytes:
    timestamp = export_timestamp()
    readme = "\n".join(
        [
            "discourseLab research export package",
            "",
            f"Generated by: discourseLab",
            f"Project: {active_project['name']}",
            f"Exported: {timestamp}",
            "",
            "Contents:",
            "- codebook.md",
            "- coded_segments.csv",
            "- coded_segments.md",
            "- memos.md",
            "- gt_hierarchy.md",
            "- cda_features.csv",
            "- voice_silence.csv",
            "- project_summary.md",
            "- project.json",
            "",
            "Uploaded source documents are not included in this package in Phase 8.",
            "",
        ]
    )
    files = {
        "codebook.md": build_codebook_markdown(active_project),
        "coded_segments.csv": generate_coded_segments_csv(active_project),
        "coded_segments.md": generate_coded_segments_markdown(active_project),
        "memos.md": generate_memos_markdown(active_project),
        "gt_hierarchy.md": generate_gt_hierarchy_markdown(active_project),
        "cda_features.csv": generate_cda_features_csv(active_project),
        "voice_silence.csv": generate_voice_silence_csv(active_project),
        "project_summary.md": generate_project_summary_markdown(active_project),
        "project.json": generate_project_json(active_project),
        "README_EXPORT.txt": readme,
    }
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, content in files.items():
            archive.writestr(filename, content.encode("utf-8"))
    return output.getvalue()


def normalize_code_color(color: str) -> str:
    color = color.strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        return color.lower()
    return DEFAULT_CODE_COLOR


def normalize_cda_color(color: str) -> str:
    color = color.strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        return color.lower()
    return DEFAULT_CDA_MARKER_COLOR


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
