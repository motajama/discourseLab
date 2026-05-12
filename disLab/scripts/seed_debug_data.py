#!/usr/bin/env python3
"""Create a broad, non-destructive sample project for local debugging."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sqlite3
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import (  # noqa: E402
    ACTOR_RELATION_TYPES,
    ACTOR_TYPES,
    CDA_MARKER_TYPES,
    DATABASE,
    DEFAULT_CDA_MARKER_COLOR,
    DISCOURSE_FEATURE_TYPES,
    MEMO_STATUSES,
    MEMO_TYPES,
    RELATION_STRENGTHS,
    RELATION_TYPES,
    UPLOAD_DIR,
)


def now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def insert(cur: sqlite3.Cursor, sql: str, params: tuple = ()) -> int:
    cur.execute(sql, params)
    return int(cur.lastrowid)


def segment_offsets(text: str, needle: str) -> tuple[int, int]:
    start = text.index(needle)
    return start, start + len(needle)


def main() -> None:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    project_name = f"Phase 11 Debug Sample {timestamp}"
    UPLOAD_DIR.mkdir(exist_ok=True)
    sample_filename = f"phase11-debug-sample-{timestamp}.txt"
    sample_text = (
        "City council members framed the housing shortage as a security issue during a public hearing. "
        "Tenant organizers challenged that framing and described evictions as a policy failure. "
        "A ministerial spokesperson promised rapid action while avoiding direct responsibility. "
        "Journalists repeated the phrase emergency measure in headlines, making the policy appear temporary. "
        "Residents in vulnerable districts described silence around rent increases and displacement."
    )
    (UPLOAD_DIR / sample_filename).write_text(sample_text, encoding="utf-8")

    con = sqlite3.connect(DATABASE)
    con.execute("PRAGMA foreign_keys = ON")
    cur = con.cursor()

    project_id = insert(
        cur,
        """
        INSERT INTO projects (
            name, description, methodology_mode, status, last_opened_at,
            research_goal, principal_investigator, created_at, updated_at
        )
        VALUES (?, ?, 'mixed', 'active', CURRENT_TIMESTAMP, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
        """,
        (
            project_name,
            "Synthetic mixed-method sample data for exercising discourseLab Phase 11.",
            "Debug how CDA framing, actor voice, and GT categories connect in analytical model exports.",
            "discourseLab debug seed",
        ),
    )

    document_ids = {}
    document_ids["hearing"] = insert(
        cur,
        """
        INSERT INTO documents (project_id, title, original_filename, file_type, text_content, note)
        VALUES (?, ?, ?, 'txt', ?, ?)
        """,
        (
            project_id,
            "Housing Hearing Excerpt",
            sample_filename,
            sample_text,
            "Seeded TXT-style source document.",
        ),
    )
    policy_text = (
        "The emergency housing package defines eligibility through risk categories. "
        "Implementation guidance foregrounds institutional coordination and backgrounds tenant testimony. "
        "The memo names consultation as evidence while giving no timetable for rent protections."
    )
    document_ids["policy"] = insert(
        cur,
        """
        INSERT INTO documents (project_id, title, original_filename, file_type, text_content, note)
        VALUES (?, ?, 'policy-brief-sample.docx', 'docx', ?, ?)
        """,
        (
            project_id,
            "Policy Brief Sample",
            policy_text,
            "Seeded DOCX-style source record without requiring a generated binary file.",
        ),
    )

    tags = {
        "housing": insert(cur, "INSERT INTO tags (project_id, name, color) VALUES (?, ?, ?)", (project_id, "housing", "#2563eb")),
        "cda": insert(cur, "INSERT INTO tags (project_id, name, color) VALUES (?, ?, ?)", (project_id, "cda", "#7c9a45")),
        "gt": insert(cur, "INSERT INTO tags (project_id, name, color) VALUES (?, ?, ?)", (project_id, "grounded theory", "#f4c542")),
    }
    cur.executemany(
        "INSERT INTO document_tags (document_id, tag_id) VALUES (?, ?)",
        [
            (document_ids["hearing"], tags["housing"]),
            (document_ids["hearing"], tags["cda"]),
            (document_ids["policy"], tags["housing"]),
            (document_ids["policy"], tags["gt"]),
        ],
    )

    segment_specs = [
        ("hearing", "Security framing", "framed the housing shortage as a security issue", "CDA: security frame."),
        ("hearing", "Tenant challenge", "Tenant organizers challenged that framing", "Negative case and counter-frame."),
        ("hearing", "Responsibility avoidance", "avoiding direct responsibility", "Agency and accountability marker."),
        ("hearing", "Emergency headline", "emergency measure in headlines", "Intertextual repetition."),
        ("hearing", "Silenced residents", "silence around rent increases and displacement", "Voice and silence example."),
        ("policy", "Risk categories", "defines eligibility through risk categories", "Policy classification."),
        ("policy", "Backgrounded testimony", "backgrounds tenant testimony", "CDA foreground/background."),
        ("policy", "Consultation evidence", "names consultation as evidence", "Evidence claim."),
    ]
    segments = {}
    for doc_key, name, selected, note in segment_specs:
        doc_text = sample_text if doc_key == "hearing" else policy_text
        start, end = segment_offsets(doc_text, selected)
        segments[name] = insert(
            cur,
            """
            INSERT INTO segments (document_id, name, selected_text, start_offset, end_offset, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (document_ids[doc_key], name, selected, start, end, note),
        )

    code_specs = [
        ("Security framing", "open", None, "#2563eb"),
        ("Responsibility avoidance", "open", None, "#dc2626"),
        ("Tenant counter-frame", "open", None, "#16a34a"),
        ("Voice and silence", "axial", None, "#7c3aed"),
        ("Policy legitimation", "axial", None, "#ea580c"),
        ("Discourse of emergency governance", "category", None, "#0891b2"),
    ]
    codes = {}
    for name, code_type, parent, color in code_specs:
        codes[name] = insert(
            cur,
            """
            INSERT INTO codes (
                project_id, name, description, code_type, color, parent_id,
                definition, include_when, exclude_when, example, analytical_note,
                gt_conditions, gt_context, gt_actions_interactions, gt_consequences,
                gt_properties, gt_dimensions, gt_theoretical_note
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                name,
                f"Debug {code_type} code for {name}.",
                code_type,
                color,
                parent,
                f"Use when data expresses {name.lower()}.",
                "Apply to explicit or strongly implied instances.",
                "Exclude vague background references.",
                f"Example for {name}.",
                "Seeded codebook note.",
                "Institutional pressure and public uncertainty.",
                "Housing policy debate.",
                "Framing, challenging, backgrounding.",
                "Legitimation or contestation of policy action.",
                "visibility, agency, responsibility",
                "low to high explicitness",
                "Links CDA features with emerging GT categories.",
            ),
        )
    cur.execute("UPDATE codes SET parent_id = ? WHERE id IN (?, ?)", (codes["Voice and silence"], codes["Tenant counter-frame"], codes["Responsibility avoidance"]))
    cur.execute("UPDATE codes SET parent_id = ? WHERE id IN (?, ?)", (codes["Policy legitimation"], codes["Security framing"], codes["Voice and silence"]))
    cur.execute("UPDATE codes SET parent_id = ? WHERE id IN (?, ?)", (codes["Discourse of emergency governance"], codes["Policy legitimation"], codes["Voice and silence"]))

    segment_code_pairs = [
        ("Security framing", "Security framing"),
        ("Tenant challenge", "Tenant counter-frame"),
        ("Responsibility avoidance", "Responsibility avoidance"),
        ("Emergency headline", "Security framing"),
        ("Silenced residents", "Voice and silence"),
        ("Risk categories", "Policy legitimation"),
        ("Backgrounded testimony", "Voice and silence"),
        ("Consultation evidence", "Policy legitimation"),
    ]
    cur.executemany(
        "INSERT INTO segment_codes (segment_id, code_id) VALUES (?, ?)",
        [(segments[segment], codes[code]) for segment, code in segment_code_pairs],
    )

    markers = {}
    for index, (marker_type, label) in enumerate(CDA_MARKER_TYPES.items()):
        markers[marker_type] = insert(
            cur,
            """
            INSERT INTO discourse_markers (project_id, name, marker_type, description, color)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                project_id,
                f"{label} marker",
                marker_type,
                f"Seeded CDA marker for {label.lower()} analysis.",
                DEFAULT_CDA_MARKER_COLOR if index % 2 else "#2563eb",
            ),
        )
    marker_links = [
        ("Security framing", "framing"),
        ("Responsibility avoidance", "agency"),
        ("Emergency headline", "intertextuality"),
        ("Silenced residents", "silence"),
        ("Backgrounded testimony", "voice"),
        ("Risk categories", "legitimation"),
    ]
    cur.executemany(
        "INSERT INTO segment_discourse_markers (segment_id, marker_id, note) VALUES (?, ?, ?)",
        [(segments[segment], markers[marker], f"Seeded marker link: {marker}") for segment, marker in marker_links],
    )

    actors = {}
    for index, (actor_type, label) in enumerate(ACTOR_TYPES.items()):
        actors[actor_type] = insert(
            cur,
            "INSERT INTO actors (project_id, name, actor_type, description) VALUES (?, ?, ?, ?)",
            (project_id, f"{label} sample actor", actor_type, f"Seeded {label.lower()} actor."),
        )
    actor_relation_keys = list(ACTOR_RELATION_TYPES)
    segment_cycle = list(segments.values())
    actor_cycle = list(actors.values())
    cur.executemany(
        "INSERT INTO segment_actors (segment_id, actor_id, relation_type, note) VALUES (?, ?, ?, ?)",
        [
            (
                segment_cycle[index % len(segment_cycle)],
                actor_cycle[index % len(actor_cycle)],
                relation_type,
                f"Seeded actor relation: {relation_type}",
            )
            for index, relation_type in enumerate(actor_relation_keys)
        ],
    )

    features = {}
    for index, (feature_type, label) in enumerate(DISCOURSE_FEATURE_TYPES.items()):
        segment_id = segment_cycle[index % len(segment_cycle)]
        features[feature_type] = insert(
            cur,
            """
            INSERT INTO discourse_features (segment_id, feature_type, value, interpretation)
            VALUES (?, ?, ?, ?)
            """,
            (
                segment_id,
                feature_type,
                f"{label} value in sample discourse",
                f"Seeded interpretation for {label.lower()}.",
            ),
        )

    memos = {}
    memo_specs = [
        ("Project orientation memo", "project", None, None, "important"),
        ("Segment evidence memo", "segment", "segment", segments["Security framing"], "use_in_article"),
        ("Code refinement memo", "code", "code", codes["Policy legitimation"], "important"),
        ("Method note", "methodological", None, None, "draft"),
        ("Theory integration memo", "theoretical", "code", codes["Discourse of emergency governance"], "use_in_article"),
        ("Reflexive note", "reflexive", None, None, "draft"),
        ("Comparison memo", "comparison", "document", document_ids["policy"], "important"),
        ("Negative case memo", "negative_case", "segment", segments["Tenant challenge"], "important"),
    ]
    for title, memo_type, linked_type, linked_id, status in memo_specs:
        if memo_type not in MEMO_TYPES or status not in MEMO_STATUSES:
            continue
        memos[title] = insert(
            cur,
            """
            INSERT INTO memos (project_id, title, body, memo_type, linked_entity_type, linked_entity_id, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                title,
                f"Seeded {memo_type} memo for debugging lists, exports, and relation maps.",
                memo_type,
                linked_type,
                linked_id,
                status,
            ),
        )

    questions = [
        insert(
            cur,
            "INSERT INTO research_questions (project_id, question, note) VALUES (?, ?, ?)",
            (project_id, "How does emergency framing legitimate housing policy?", "Connect CDA framing with GT category."),
        ),
        insert(
            cur,
            "INSERT INTO research_questions (project_id, question, note) VALUES (?, ?, ?)",
            (project_id, "Which actors are foregrounded or backgrounded in the discourse?", "Useful for CDA actor reports."),
        ),
    ]

    relation_entities = [
        ("segment", segments["Security framing"], "code", codes["Security framing"]),
        ("segment", segments["Tenant challenge"], "code", codes["Tenant counter-frame"]),
        ("code", codes["Security framing"], "code", codes["Policy legitimation"]),
        ("code", codes["Policy legitimation"], "code", codes["Discourse of emergency governance"]),
        ("actor", actors["politician"], "code", codes["Security framing"]),
        ("actor", actors["vulnerable_group"], "discourse_feature", features["agency"]),
        ("discourse_marker", markers["legitimation"], "code", codes["Policy legitimation"]),
        ("discourse_feature", features["framing"], "research_question", questions[0]),
        ("memo", memos["Theory integration memo"], "code", codes["Discourse of emergency governance"]),
        ("document", document_ids["policy"], "segment", segments["Risk categories"]),
        ("research_question", questions[1], "actor", actors["journalist"]),
    ]
    relation_type_keys = list(RELATION_TYPES)
    strength_keys = list(RELATION_STRENGTHS)
    for index, relation_type in enumerate(relation_type_keys):
        source_type, source_id, target_type, target_id = relation_entities[index % len(relation_entities)]
        strength = strength_keys[index % len(strength_keys)]
        insert(
            cur,
            """
            INSERT INTO relations (
                project_id, source_type, source_id, target_type, target_id,
                relation_type, title, strength, memo, evidence_note, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                source_type,
                source_id,
                target_type,
                target_id,
                relation_type,
                f"Debug relation: {relation_type}",
                strength,
                f"Seeded analytical memo for {relation_type}.",
                "Synthetic evidence note for export and model debugging.",
                now(),
            ),
        )

    audit_rows = [
        (project_id, "project", project_id, "seed_debug_data", f"Created {project_name}."),
        (project_id, "document", document_ids["hearing"], "import_document", "Seeded hearing document."),
        (project_id, "segment", segments["Security framing"], "create_segment", "Seeded named segment."),
        (project_id, "code", codes["Security framing"], "create_code", "Seeded codebook entry."),
        (project_id, "relation", None, "create_relation", "Seeded broad analytical relation set."),
        (project_id, "export", None, "debug_ready", "Project covers Phase 11 export paths."),
    ]
    cur.executemany(
        "INSERT INTO audit_log (project_id, entity_type, entity_id, action, details) VALUES (?, ?, ?, ?, ?)",
        audit_rows,
    )

    con.commit()
    con.close()
    print(f"Created sample project {project_id}: {project_name}")
    print(f"Sample source file: uploads/{sample_filename}")
    print("Open /projects and choose the new sample project if it is not already active in your browser session.")


if __name__ == "__main__":
    main()
