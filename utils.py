# =========================================================
# 選考履歴からステータスを判定
# =========================================================

def determine_status(event_type, title):
    """
    選考履歴の種類とタイトルから、
    企業の現在ステータスを判定する。
    """

    event_type = (event_type or "").strip()
    title = (title or "").strip()

    combined_text = f"{event_type} {title}".lower()

    # -----------------------------------------------------
    # 結果関係
    # -----------------------------------------------------

    if (
        "不合格" in combined_text
        or "お祈り" in combined_text
    ):
        return "不合格"

    if (
        "内定" in combined_text
        or "合格" in combined_text
    ):
        return "合格"

    if (
        "結果待ち" in combined_text
        or event_type == "結果通知"
    ):
        return "結果待ち"

    # -----------------------------------------------------
    # 面接関係
    # -----------------------------------------------------

    if "最終面接" in combined_text:
        return "最終面接"

    if "三次面接" in combined_text:
        return "三次面接"

    if "二次面接" in combined_text:
        return "二次面接"

    if "一次面接" in combined_text:
        return "一次面接"

    if "面接" in combined_text:
        return "面接"

    # -----------------------------------------------------
    # 適性検査関係
    # -----------------------------------------------------

    if (
        "適性検査" in combined_text
        or "spi" in combined_text
        or "webテスト" in combined_text
        or "web test" in combined_text
        or "玉手箱" in combined_text
        or "テストセンター" in combined_text
    ):
        return "適性検査"

    # -----------------------------------------------------
    # ES関係
    # -----------------------------------------------------

    if (
        "es提出" in combined_text
        or "エントリーシート提出" in combined_text
    ):
        return "ES提出済"

    if (
        "es作成" in combined_text
        or "エントリーシート作成" in combined_text
    ):
        return "ES作成中"

    if event_type == "ES":
        return "ES提出済"

    # -----------------------------------------------------
    # その他
    # -----------------------------------------------------

    if event_type == "インターン":
        return "インターン"

    if event_type == "説明会":
        return "説明会"

    # 自動判定できない場合
    return None


# =========================================================
# 選考履歴から企業ステータスを再計算
# =========================================================

def update_company_status_from_events(
    connection,
    company_id,
):
    """
    残っている最新の選考履歴から、
    企業ステータスを再計算する。
    """

    events = connection.execute(
        """
        SELECT
            event_type,
            title
        FROM company_events
        WHERE company_id = ?
        ORDER BY id DESC
        """,
        (company_id,),
    ).fetchall()

    for event in events:
        status = determine_status(
            event["event_type"],
            event["title"],
        )

        if status:
            connection.execute(
                """
                UPDATE companies
                SET status = ?
                WHERE id = ?
                """,
                (
                    status,
                    company_id,
                ),
            )

            return

    # 判定可能な履歴が1件もない場合は未応募に戻す
    connection.execute(
        """
        UPDATE companies
        SET status = ?
        WHERE id = ?
        """,
        (
            "未応募",
            company_id,
        ),
    )