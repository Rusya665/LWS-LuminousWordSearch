"""
lws_engine.database
===================
SQLite database manager with FTS5 full-text indexing, WAL mode, document bundling,
figure metadata storage, and user annotations.
Thread-safe with RLock and check_same_thread=False.
"""

from typing import Dict, Any, List, Optional, Tuple
import sqlite3
import os
import json
import datetime
import threading
import logging
import re
from .config import win_safe_path

logger = logging.getLogger(__name__)


class LwsDatabase:
    """Manages the SQLite FTS5 database for documents, sentences, and figures."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        if self.db_path != ":memory:":
            os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        self._lock = threading.RLock()
        self._conn: Optional[sqlite3.Connection] = None
        self._init_db()

    def get_connection(self) -> sqlite3.Connection:
        with self._lock:
            if self._conn is None:
                target_path = ":memory:" if self.db_path == ":memory:" else win_safe_path(self.db_path)
                conn = sqlite3.connect(target_path, timeout=30.0, check_same_thread=False)
                conn.row_factory = sqlite3.Row
                if self.db_path != ":memory:":
                    conn.execute("PRAGMA journal_mode = WAL;")
                conn.execute("PRAGMA synchronous = NORMAL;")
                conn.execute("PRAGMA foreign_keys = ON;")
                self._conn = conn
            return self._conn

    def close(self):
        """Closes any open database connections."""
        with self._lock:
            if self._conn is not None:
                try:
                    self._conn.close()
                except Exception as e:
                    logger.debug(f"Error closing DB connection: {e}")
                self._conn = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def _init_db(self):
        with self._lock:
            conn = self.get_connection()
            with conn:
                conn.executescript("""
                    CREATE TABLE IF NOT EXISTS documents (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        watch_folder TEXT NOT NULL,
                        rel_path TEXT NOT NULL UNIQUE,
                        filename TEXT NOT NULL,
                        title TEXT,
                        authors TEXT,
                        year INTEGER,
                        doi TEXT,
                        doc_role TEXT DEFAULT 'MAIN',
                        bundle_id TEXT,
                        page_count INTEGER DEFAULT 0,
                        file_size INTEGER DEFAULT 0,
                        mtime REAL DEFAULT 0,
                        md_path TEXT,
                        abstract TEXT
                    );
                    CREATE INDEX IF NOT EXISTS idx_docs_bundle ON documents(bundle_id);
                    CREATE INDEX IF NOT EXISTS idx_docs_role ON documents(doc_role);
                    CREATE INDEX IF NOT EXISTS idx_docs_rel_path ON documents(rel_path);

                    CREATE VIRTUAL TABLE IF NOT EXISTS sentences_fts USING fts5(
                        doc_id UNINDEXED,
                        page_number UNINDEXED,
                        sentence_index UNINDEXED,
                        sentence_text,
                        tokenize = 'unicode61 remove_diacritics 2'
                    );

                    CREATE TABLE IF NOT EXISTS figures (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        doc_id INTEGER NOT NULL,
                        figure_index INTEGER NOT NULL,
                        page_number INTEGER NOT NULL,
                        rel_image_path TEXT NOT NULL,
                        caption TEXT,
                        width INTEGER DEFAULT 0,
                        height INTEGER DEFAULT 0,
                        ocr_text TEXT,
                        FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE
                    );
                    CREATE INDEX IF NOT EXISTS idx_figures_doc ON figures(doc_id);

                    CREATE VIRTUAL TABLE IF NOT EXISTS figures_fts USING fts5(
                        figure_id UNINDEXED,
                        doc_id UNINDEXED,
                        caption,
                        ocr_text,
                        tokenize = 'unicode61 remove_diacritics 2'
                    );

                    CREATE TABLE IF NOT EXISTS user_annotations (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        doc_id INTEGER NOT NULL,
                        page_number INTEGER,
                        sentence_index INTEGER,
                        note TEXT NOT NULL,
                        tags TEXT,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE
                    );
                    CREATE INDEX IF NOT EXISTS idx_annotations_doc ON user_annotations(doc_id);
                """)

    def upsert_document(self, doc_data: Dict[str, Any]) -> int:
        """Inserts or updates a document record, returning the document id."""
        sql = """
            INSERT INTO documents (
                watch_folder, rel_path, filename, title, authors, year, doi,
                doc_role, bundle_id, page_count, file_size, mtime, md_path, abstract
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(rel_path) DO UPDATE SET
                watch_folder = excluded.watch_folder,
                filename = excluded.filename,
                title = excluded.title,
                authors = excluded.authors,
                year = excluded.year,
                doi = excluded.doi,
                doc_role = excluded.doc_role,
                bundle_id = excluded.bundle_id,
                page_count = excluded.page_count,
                file_size = excluded.file_size,
                mtime = excluded.mtime,
                md_path = excluded.md_path,
                abstract = excluded.abstract
        """
        params = (
            doc_data.get("watch_folder", ""),
            doc_data.get("rel_path", ""),
            doc_data.get("filename", ""),
            doc_data.get("title", ""),
            doc_data.get("authors", ""),
            doc_data.get("year", 0),
            doc_data.get("doi", ""),
            doc_data.get("doc_role", "MAIN"),
            doc_data.get("bundle_id", ""),
            doc_data.get("page_count", 0),
            doc_data.get("file_size", 0),
            doc_data.get("mtime", 0.0),
            doc_data.get("md_path", ""),
            doc_data.get("abstract", "")
        )
        with self._lock:
            conn = self.get_connection()
            with conn:
                cur = conn.cursor()
                cur.execute(sql, params)
                cur.execute("SELECT id FROM documents WHERE rel_path = ?", (doc_data.get("rel_path", ""),))
                row = cur.fetchone()
                return row["id"] if row else 0

    def clear_document_content(self, doc_id: int):
        """Clears sentences and figures for a document prior to re-indexing."""
        with self._lock:
            conn = self.get_connection()
            with conn:
                conn.execute("DELETE FROM sentences_fts WHERE doc_id = ?", (str(doc_id),))
                cur = conn.cursor()
                cur.execute("SELECT id FROM figures WHERE doc_id = ?", (doc_id,))
                fig_ids = [str(r["id"]) for r in cur.fetchall()]
                for fid in fig_ids:
                    conn.execute("DELETE FROM figures_fts WHERE figure_id = ?", (fid,))
                conn.execute("DELETE FROM figures_fts WHERE doc_id = ?", (str(doc_id),))
                conn.execute("DELETE FROM figures WHERE doc_id = ?", (doc_id,))

    def delete_document_by_rel_path(self, rel_path: str) -> bool:
        """
        Atomically removes a document and its sentences, figures,
        FTS entries, and user annotations by its relative path.
        Returns True if a document was deleted, False if not found.
        """
        if not rel_path:
            return False
        with self._lock:
            conn = self.get_connection()
            with conn:
                cur = conn.cursor()
                norm_p = rel_path.replace("\\", "/")
                cur.execute(
                    "SELECT id FROM documents WHERE rel_path = ? OR rel_path = ?",
                    (rel_path, norm_p)
                )
                row = cur.fetchone()
                if not row:
                    return False
                doc_id = row["id"]

                conn.execute("DELETE FROM sentences_fts WHERE doc_id = ?", (str(doc_id),))
                cur.execute("SELECT id FROM figures WHERE doc_id = ?", (doc_id,))
                fig_ids = [str(r["id"]) for r in cur.fetchall()]
                for fid in fig_ids:
                    conn.execute("DELETE FROM figures_fts WHERE figure_id = ?", (fid,))
                conn.execute("DELETE FROM figures_fts WHERE doc_id = ?", (str(doc_id),))
                conn.execute("DELETE FROM figures WHERE doc_id = ?", (doc_id,))
                conn.execute("DELETE FROM user_annotations WHERE doc_id = ?", (doc_id,))
                conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
                return True

    def index_sentences(self, doc_id: int, sentences: List[Tuple[int, int, str]]):
        """
        Inserts sentences into sentences_fts.
        sentences is a list of tuples: (page_number, sentence_index, sentence_text)
        """
        with self._lock:
            conn = self.get_connection()
            with conn:
                conn.executemany(
                    "INSERT INTO sentences_fts(doc_id, page_number, sentence_index, sentence_text) VALUES (?, ?, ?, ?)",
                    [(str(doc_id), str(p), str(idx), text) for p, idx, text in sentences]
                )

    def index_figures(self, doc_id: int, figures: List[Dict[str, Any]]):
        """Inserts extracted figures and their captions into the DB and figures_fts."""
        with self._lock:
            conn = self.get_connection()
            with conn:
                cur = conn.cursor()
                for fig in figures:
                    fig_idx_raw = fig.get("figure_index") if fig.get("figure_index") is not None else fig.get("fig_index", 0)
                    page_num_raw = fig.get("page_number") if fig.get("page_number") is not None else fig.get("page_num", 1)

                    try:
                        fig_idx = int(fig_idx_raw) if fig_idx_raw is not None else 0
                    except (ValueError, TypeError):
                        fig_idx = 0

                    try:
                        page_num = int(page_num_raw) if page_num_raw is not None else 1
                    except (ValueError, TypeError):
                        page_num = 1

                    rel_p = str(fig.get("rel_image_path") or fig.get("file_path") or fig.get("image_path") or "").replace("\\", "/")
                    caption_str = str(fig.get("caption") or "")

                    try:
                        width_val = int(fig.get("width") or 0)
                    except (ValueError, TypeError):
                        width_val = 0

                    try:
                        height_val = int(fig.get("height") or 0)
                    except (ValueError, TypeError):
                        height_val = 0

                    ocr_str = str(fig.get("ocr_text") or "")

                    cur.execute("""
                        INSERT INTO figures (doc_id, figure_index, page_number, rel_image_path, caption, width, height, ocr_text)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        doc_id,
                        fig_idx,
                        page_num,
                        rel_p,
                        caption_str,
                        width_val,
                        height_val,
                        ocr_str
                    ))
                    fig_id = cur.lastrowid
                    cur.execute("""
                        INSERT INTO figures_fts (figure_id, doc_id, caption, ocr_text)
                        VALUES (?, ?, ?, ?)
                    """, (
                        str(fig_id),
                        str(doc_id),
                        caption_str,
                        ocr_str
                    ))

    def _build_fts_query(self, query: str, synonyms: Optional[List[str]] = None) -> str:
        """
        Parses and formats user query into a safe, valid SQLite FTS5 query string.
        Handles explicit boolean keywords (AND, OR, NOT), phrases, bare/hyphenated words,
        and synonyms without breaking on hyphens or syntax errors.
        """
        if not query or not query.strip():
            return ""

        q = query.strip()

        # Check for explicit boolean operators
        has_bool_ops = bool(re.search(r'\b(?:AND|OR|NOT)\b', q, flags=re.IGNORECASE))

        # Balance odd quotes
        if q.count('"') % 2 != 0:
            q = q.replace('"', ' ')

        if has_bool_ops:
            # Tokenize: quoted strings, parens, boolean keywords, or bare words
            token_re = re.compile(r'"([^"]*)"|(\(|\))|(\b(?:AND|OR|NOT)\b)|([^\s()"]+)', flags=re.IGNORECASE)
            raw_tokens = []
            for m in token_re.finditer(q):
                quoted, paren, op, word = m.groups()
                if quoted is not None:
                    cleaned_val = quoted.strip().replace('"', '""')
                    if cleaned_val:
                        raw_tokens.append(('TERM', f'"{cleaned_val}"'))
                elif paren:
                    raw_tokens.append(('PAREN', paren))
                elif op:
                    raw_tokens.append(('OP', op.upper()))
                elif word:
                    cleaned_word = word.strip().replace('"', '""')
                    if cleaned_word:
                        raw_tokens.append(('TERM', f'"{cleaned_word}"'))

            # Clean tokens to avoid FTS5 syntax errors
            valid_tokens = []
            expect_term = True

            for kind, val in raw_tokens:
                if kind == 'TERM':
                    if not expect_term:
                        valid_tokens.append('AND')
                    valid_tokens.append(val)
                    expect_term = False
                elif kind == 'PAREN':
                    valid_tokens.append(val)
                    if val == '(':
                        expect_term = True
                    else:
                        expect_term = False
                elif kind == 'OP':
                    if val in ('AND', 'OR'):
                        if not expect_term:
                            valid_tokens.append(val)
                            expect_term = True
                    elif val == 'NOT':
                        valid_tokens.append(val)
                        expect_term = True

            # Drop any trailing operators or open parens
            while valid_tokens and valid_tokens[-1] in ('AND', 'OR', 'NOT', '('):
                valid_tokens.pop()

            if not valid_tokens:
                words = re.findall(r'[a-zA-Z0-9_\-]+', q)
                return " AND ".join(f'"{w}"' for w in words) if words else ""

            return " ".join(valid_tokens)

        else:
            # Plain single or multi-word query without boolean operators
            quoted_phrases = re.findall(r'"([^"]+)"', q)
            clean_no_quotes = re.sub(r'"[^"]*"', ' ', q)
            words = [w.strip().replace('"', '""') for w in re.findall(r'[^\s()"]+', clean_no_quotes) if w.strip()]

            all_terms = []
            for p in quoted_phrases:
                p_clean = p.strip().replace('"', '""')
                if p_clean:
                    all_terms.append(f'"{p_clean}"')
            for w in words:
                all_terms.append(f'"{w}"')

            if not all_terms:
                return ""

            if len(all_terms) == 1:
                main_clause = all_terms[0]
            else:
                exact_phrase = f'"{q.strip().replace(chr(34), "")}"'
                words_and = " AND ".join(all_terms)
                main_clause = f'({exact_phrase} OR ({words_and}))'

            if synonyms:
                syn_clauses = []
                for s in synonyms:
                    s_clean = s.strip().replace('"', '""')
                    if s_clean and f'"{s_clean}"' != main_clause:
                        syn_clauses.append(f'"{s_clean}"')
                if syn_clauses:
                    main_clause = f'({main_clause} OR ' + " OR ".join(syn_clauses) + ')'

            return main_clause

    def search_fts(
        self,
        query: str,
        synonyms: Optional[List[str]] = None,
        role_filter: Optional[str] = None,
        bundle_id: Optional[str] = None,
        limit: int = 60,
        offset: int = 0
    ) -> List[Dict[str, Any]]:
        """
        Executes an FTS5 MATCH query on sentences.
        Constructs boolean search with provided synonyms and operators.
        Returns matched sentences with parent document metadata and highlight snippet.
        """
        if not query or not query.strip():
            return []

        fts_query = self._build_fts_query(query, synonyms=synonyms)
        if not fts_query:
            return []

        sql = """
            SELECT 
                s.doc_id,
                s.page_number,
                s.sentence_index,
                s.sentence_text,
                snippet(sentences_fts, 3, '<mark class="hit-direct">', '</mark>', '...', 32) AS highlighted_snippet,
                d.filename,
                d.title,
                d.authors,
                d.year,
                d.doi,
                d.doc_role,
                d.bundle_id,
                d.rel_path,
                d.watch_folder
            FROM sentences_fts s
            JOIN documents d ON CAST(s.doc_id AS INTEGER) = d.id
            WHERE sentences_fts MATCH ?
        """
        params: List[Any] = [fts_query]

        if role_filter and role_filter != "ALL":
            sql += " AND d.doc_role = ?"
            params.append(role_filter)

        if bundle_id:
            sql += " AND d.bundle_id = ?"
            params.append(bundle_id)

        sql += " ORDER BY rank LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            results = []
            try:
                cur.execute(sql, params)
                for r in cur.fetchall():
                    results.append(dict(r))
            except sqlite3.OperationalError as e:
                logger.debug(f"FTS query fallback for '{query}': {e}")
                try:
                    words = re.findall(r'[a-zA-Z0-9_\-]+', query)
                    fallback_query = " AND ".join(f'"{w}"' for w in words) if words else ""
                    if fallback_query:
                        params[0] = fallback_query
                        cur.execute(sql, params)
                        for r in cur.fetchall():
                            results.append(dict(r))
                except Exception as e2:
                    logger.debug(f"FTS secondary fallback failed for '{query}': {e2}")

            return results

    def search_regex(
        self,
        pattern: re.Pattern,
        role_filter: Optional[str] = None,
        bundle_id: Optional[str] = None,
        limit: int = 80
    ) -> List[Dict[str, Any]]:
        """
        Scans sentences using a compiled regular expression,
        highlighting matches with <mark class="hit-direct">.
        """
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            sql = """
                SELECT 
                    s.doc_id,
                    s.page_number,
                    s.sentence_index,
                    s.sentence_text,
                    d.filename,
                    d.title,
                    d.authors,
                    d.year,
                    d.doi,
                    d.doc_role,
                    d.bundle_id,
                    d.rel_path,
                    d.watch_folder
                FROM sentences_fts s
                JOIN documents d ON CAST(s.doc_id AS INTEGER) = d.id
            """
            params: List[Any] = []
            conditions = []
            if role_filter and role_filter != "ALL":
                conditions.append("d.doc_role = ?")
                params.append(role_filter)
            if bundle_id:
                conditions.append("d.bundle_id = ?")
                params.append(bundle_id)
            if conditions:
                sql += " WHERE " + " AND ".join(conditions)

            cur.execute(sql, params)
            results = []
            for row in cur:
                text = row["sentence_text"]
                if pattern.search(text):
                    highlighted = pattern.sub(r'<mark class="hit-direct">\g<0></mark>', text)
                    item = dict(row)
                    item["highlighted_snippet"] = highlighted
                    results.append(item)
                    if len(results) >= limit:
                        break
            return results

    def search_figures(self, query: str, limit: int = 40) -> List[Dict[str, Any]]:
        """Searches figure captions and OCR text using FTS5."""
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            if not query or not query.strip():
                cur.execute("""
                    SELECT f.*, d.filename, d.title, d.bundle_id, d.watch_folder, d.rel_path
                    FROM figures f
                    JOIN documents d ON f.doc_id = d.id
                    ORDER BY f.id DESC LIMIT ?
                """, (limit,))
                return [dict(r) for r in cur.fetchall()]

            fts_query = self._build_fts_query(query)
            if not fts_query:
                return []

            sql = """
                SELECT f.*, d.filename, d.title, d.bundle_id, d.watch_folder, d.rel_path,
                       snippet(figures_fts, 2, '<mark>', '</mark>', '...', 25) AS caption_snippet
                FROM figures_fts ff
                JOIN figures f ON CAST(ff.figure_id AS INTEGER) = f.id
                JOIN documents d ON f.doc_id = d.id
                WHERE figures_fts MATCH ?
                ORDER BY rank LIMIT ?
            """
            results = []
            try:
                cur.execute(sql, (fts_query, limit))
                for r in cur.fetchall():
                    results.append(dict(r))
            except sqlite3.OperationalError as e:
                logger.debug(f"Figures FTS fallback for '{query}': {e}")
                try:
                    words = re.findall(r'[a-zA-Z0-9_\-]+', query)
                    fallback_query = " AND ".join(f'"{w}"' for w in words) if words else ""
                    if fallback_query:
                        cur.execute(sql, (fallback_query, limit))
                        for r in cur.fetchall():
                            results.append(dict(r))
                except Exception as e2:
                    logger.debug(f"Figures FTS secondary fallback failed: {e2}")
            return results

    def get_bundle(self, bundle_id: str) -> List[Dict[str, Any]]:
        """Returns all documents associated with a bundle_id (Main + SI + Peer Review)."""
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("""
                SELECT * FROM documents 
                WHERE bundle_id = ? 
                ORDER BY CASE doc_role 
                    WHEN 'MAIN' THEN 1 
                    WHEN 'SI' THEN 2 
                    WHEN 'PEER_REVIEW' THEN 3 
                    WHEN 'PATENT' THEN 4 
                    ELSE 5 END, filename ASC
            """, (bundle_id,))
            return [dict(r) for r in cur.fetchall()]

    def get_document_by_id(self, doc_id: int) -> Optional[Dict[str, Any]]:
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("SELECT * FROM documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
            return dict(row) if row else None

    def get_figures_for_doc(self, doc_id: int) -> List[Dict[str, Any]]:
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("SELECT * FROM figures WHERE doc_id = ? ORDER BY page_number, figure_index", (doc_id,))
            return [dict(r) for r in cur.fetchall()]

    def list_all_documents(self, limit: int = 500) -> List[Dict[str, Any]]:
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("""
                SELECT id, filename, title, authors, year, doi, doc_role, bundle_id, rel_path, watch_folder, page_count
                FROM documents
                ORDER BY year DESC, filename ASC
                LIMIT ?
            """, (limit,))
            return [dict(r) for r in cur.fetchall()]

    def save_annotation(self, doc_id: int, page_number: int, sentence_index: int, note: str, tags: str = "") -> int:
        now_iso = datetime.datetime.now().isoformat()
        with self._lock:
            conn = self.get_connection()
            with conn:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO user_annotations (doc_id, page_number, sentence_index, note, tags, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (doc_id, page_number, sentence_index, note, tags, now_iso, now_iso))
                return cur.lastrowid

    def get_annotations(self, doc_id: int) -> List[Dict[str, Any]]:
        with self._lock:
            conn = self.get_connection()
            cur = conn.cursor()
            cur.execute("SELECT * FROM user_annotations WHERE doc_id = ? ORDER BY page_number, sentence_index", (doc_id,))
            return [dict(r) for r in cur.fetchall()]
