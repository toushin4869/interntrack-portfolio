from datetime import date, timedelta

from flask import Blueprint, render_template

from database import get_db_connection


dashboard_bp = Blueprint(
    "dashboard",
    __name__,
)


# =========================================================
# ホーム画面
# =========================================================

@dashboard_bp.route("/")
def index():
    connection = get_db_connection()

    today = date.today()
    seven_days_later = today + timedelta(days=7)

    try:
        # -------------------------------------------------
        # 登録企業数
        # -------------------------------------------------

        company_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM companies
            """
        ).fetchone()["count"]

        # -------------------------------------------------
        # 選考中の企業数
        # -------------------------------------------------

        selection_statuses = (
            "ES作成中",
            "ES提出済",
            "適性検査",
            "面接",
            "一次面接",
            "二次面接",
            "三次面接",
            "最終面接",
        )

        placeholders = ",".join(
            "?"
            for _ in selection_statuses
        )

        selection_count = connection.execute(
            f"""
            SELECT COUNT(*) AS count
            FROM companies
            WHERE status IN ({placeholders})
            """,
            selection_statuses,
        ).fetchone()["count"]

        # -------------------------------------------------
        # 7日以内の締切件数
        # -------------------------------------------------

        deadline_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM companies
            WHERE deadline IS NOT NULL
              AND deadline != ''
              AND date(deadline)
                  BETWEEN date(?) AND date(?)
            """,
            (
                today.isoformat(),
                seven_days_later.isoformat(),
            ),
        ).fetchone()["count"]

        # -------------------------------------------------
        # 結果待ち企業数
        # -------------------------------------------------

        waiting_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM companies
            WHERE status = ?
            """,
            ("結果待ち",),
        ).fetchone()["count"]

        # -------------------------------------------------
        # 今日の予定
        # -------------------------------------------------

        today_events = connection.execute(
            """
            SELECT
                company_events.*,
                companies.name AS company_name
            FROM company_events
            JOIN companies
              ON companies.id = company_events.company_id
            WHERE date(company_events.event_date) = date(?)
            ORDER BY
                company_events.event_date ASC,
                company_events.id ASC
            """,
            (
                today.isoformat(),
            ),
        ).fetchall()

        # -------------------------------------------------
        # 明日から7日以内の予定
        # -------------------------------------------------

        upcoming_events = connection.execute(
            """
            SELECT
                company_events.*,
                companies.name AS company_name
            FROM company_events
            JOIN companies
              ON companies.id = company_events.company_id
            WHERE date(company_events.event_date)
                BETWEEN date(?, '+1 day') AND date(?)
            ORDER BY
                company_events.event_date ASC,
                company_events.id ASC
            """,
            (
                today.isoformat(),
                seven_days_later.isoformat(),
            ),
        ).fetchall()

        # -------------------------------------------------
        # 7日以内の締切
        # -------------------------------------------------

        upcoming_deadlines = connection.execute(
            """
            SELECT *
            FROM companies
            WHERE deadline IS NOT NULL
              AND deadline != ''
              AND date(deadline)
                  BETWEEN date(?) AND date(?)
            ORDER BY
                deadline ASC,
                id ASC
            """,
            (
                today.isoformat(),
                seven_days_later.isoformat(),
            ),
        ).fetchall()

    finally:
        connection.close()

    return render_template(
        "index.html",
        company_count=company_count,
        selection_count=selection_count,
        deadline_count=deadline_count,
        waiting_count=waiting_count,
        today_events=today_events,
        upcoming_events=upcoming_events,
        upcoming_deadlines=upcoming_deadlines,
    )