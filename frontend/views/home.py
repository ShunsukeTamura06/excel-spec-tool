"""ホーム — ジョブ管理 + 新規アップロード.

アップロード時に extract と analyze を 1 ボタンで自動実行し、
進捗を `st.status` でフェーズ表示する。
"""

from __future__ import annotations

import streamlit as st

from frontend._sidebar import render_job_sidebar
from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def _set_current_job(job_id: str) -> None:
    st.session_state["job_id"] = job_id


def _do_analysis(client: BackendClient, filename: str, data: bytes) -> str | None:
    """Upload → extract → analyze を 1 操作で完了させる.

    `st.status` で 2 段階のフェーズを順に表示する。
    """
    with st.status("Excel ファイルを分析中...", expanded=True) as status:
        try:
            st.write("📥 **ステップ 1/2:** ファイル送信・抽出")
            st.caption(
                "VBA・数式・参照関係・全セルデータを抽出します。"
                "大きなファイルだと数十秒かかります。"
            )
            job_id = client.extract(filename, data)
            st.write(f"✅ 抽出完了 (`{job_id[:8]}...`)")

            st.write("📝 **ステップ 2/2:** 設計書生成")
            client.analyze(job_id)
            st.write("✅ 設計書生成完了")

            status.update(label="✅ 分析完了", state="complete", expanded=False)
            return job_id
        except BackendError as e:
            status.update(label=f"❌ 失敗: {e.detail}", state="error", expanded=True)
            st.error(f"Backend エラー ({e.status_code}): {e.detail}")
            return None


# ========== ページ本体 ==========

st.title("🏠 ホーム")
st.caption(
    "VBA / 数式 / 参照関係を含む `.xlsm` `.xls` ツールの統合設計書を生成し、改修対話を支援します。"
)

client = _client()
if not client.health():
    st.error(
        f"Backend ({client.base_url}) に接続できません。"
        " `uv run uvicorn backend.main:app --reload --port 8000` を起動してください。"
    )
    st.stop()

# ホームではサイドバーの「ホームに戻る」ボタンは出さない
render_job_sidebar(show_home_button=False)

# ===== 現在のジョブ =====
current_job_id = st.session_state.get("job_id")
current_job: dict | None = None
if current_job_id:
    try:
        jobs_list = client.list_jobs()
        current_job = next((j for j in jobs_list if j["job_id"] == current_job_id), None)
    except BackendError:
        pass

if current_job:
    st.success(
        f"📂 **現在のジョブ:** {current_job['filename']}"
        f" ・ 状態: `{current_job['status']}`"
        f" ・ 作成: {current_job['created_at'][:19]}"
    )
    cols = st.columns([2, 2, 2, 4])
    with cols[0]:
        if st.button("📑 設計書を見る", use_container_width=True):
            st.switch_page("views/spec.py")
    with cols[1]:
        if st.button("💬 チャットする", use_container_width=True):
            st.switch_page("views/chat.py")
    with cols[2]:
        if st.button("🗑️ 削除", use_container_width=True):
            try:
                client.delete_job(current_job_id)
                st.session_state.pop("job_id", None)
                st.rerun()
            except BackendError as e:
                st.error(f"削除失敗: {e}")

    st.divider()

# ===== 新規アップロード =====
st.subheader("📤 別の Excel を分析する" if current_job else "📤 Excel を分析する")
uploaded = st.file_uploader(
    "ファイルを選択 (.xlsm / .xls / .xlsx)",
    type=["xlsm", "xls", "xlsx"],
    accept_multiple_files=False,
)
if uploaded is not None and st.button("分析を開始", type="primary"):
    new_job_id = _do_analysis(client, uploaded.name, uploaded.getvalue())
    if new_job_id:
        _set_current_job(new_job_id)
        st.rerun()

st.divider()

# ===== 過去のジョブ =====
st.subheader("📁 過去のジョブから開く")
try:
    jobs = client.list_jobs()
except BackendError as e:
    st.error(f"ジョブ一覧取得失敗: {e}")
    st.stop()

if not jobs:
    st.info("過去のジョブはありません。上のフォームから新規にアップロードしてください。")
else:
    for job in jobs:
        is_current = job["job_id"] == current_job_id
        with st.container(border=True):
            cols = st.columns([4, 2, 2, 1, 1])
            with cols[0]:
                badge = "  🔵 選択中" if is_current else ""
                st.write(f"**{job['filename']}**{badge}")
                st.caption(f"`{job['job_id']}`")
            with cols[1]:
                st.caption("作成日時")
                st.write(job["created_at"][:19])
            with cols[2]:
                st.caption("状態")
                st.write(f"`{job['status']}`")
            with cols[3]:
                if is_current:
                    st.button(
                        "選択中",
                        disabled=True,
                        use_container_width=True,
                        key=f"current_{job['job_id']}",
                    )
                else:
                    if st.button("開く", key=f"open_{job['job_id']}", use_container_width=True):
                        _set_current_job(job["job_id"])
                        st.rerun()
            with cols[4]:
                if st.button("削除", key=f"del_{job['job_id']}", use_container_width=True):
                    try:
                        client.delete_job(job["job_id"])
                        if is_current:
                            st.session_state.pop("job_id", None)
                        st.rerun()
                    except BackendError as e:
                        st.error(f"削除失敗: {e}")
