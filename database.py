import sqlite3

from config import DATABASE


# =========================================================
# データベース接続
# =========================================================

def get_db_connection():
    """
    SQLiteデータベースへ接続する。

    sqlite3.Rowを設定することで、
    データを row["name"] の形式で取得できるようにする。
    """

    connection = sqlite3.connect(DATABASE)

    connection.row_factory = sqlite3.Row

    # 外部キー制約を有効にする
    connection.execute(
        "PRAGMA foreign_keys = ON"
    )

    return connection


# =========================================================
# テーブル作成
# =========================================================

def create_tables():
    """
    InternTrackで使用するテーブルを作成する。

    すでにテーブルが存在する場合は、
    既存のテーブルをそのまま使用する。
    """

    connection = get_db_connection()

    try:
        # -------------------------------------------------
        # 企業情報
        # -------------------------------------------------

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                status TEXT NOT NULL,
                next_schedule TEXT,
                deadline TEXT,
                mypage_url TEXT,
                memo TEXT
            )
            """
        )

        # -------------------------------------------------
        # 選考履歴
        # -------------------------------------------------

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS company_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_id INTEGER NOT NULL,
                event_date TEXT,
                event_type TEXT NOT NULL,
                title TEXT NOT NULL,
                memo TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,

                FOREIGN KEY (company_id)
                    REFERENCES companies (id)
                    ON DELETE CASCADE
            )
            """
        )

        # -------------------------------------------------
        # 企業メモ
        # -------------------------------------------------

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS company_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,

                FOREIGN KEY (company_id)
                    REFERENCES companies (id)
                    ON DELETE CASCADE
            )
            """
        )

        # -------------------------------------------------
        # 古いデータベースへの対応
        # -------------------------------------------------

        columns = connection.execute(
            """
            PRAGMA table_info(companies)
            """
        ).fetchall()

        column_names = [
            column["name"]
            for column in columns
        ]

        # 以前作成したDBにmemo列がない場合だけ追加する
        if "memo" not in column_names:
            connection.execute(
                """
                ALTER TABLE companies
                ADD COLUMN memo TEXT
                """
            )

        connection.commit()

    finally:
        connection.close()