"""Streamlit エントリ. トップページ.

ファイルアップロード + 過去ジョブ一覧 + ジョブ削除を提供する。
SPEC.md §6.1 / §6.2 参照。

起動例:
    uv run streamlit run frontend/app.py --server.port 8501
"""

from __future__ import annotations

import streamlit as st

from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    """Streamlit のセッションごとに 1 つの BackendClient を確保する."""
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def _set_current_job(job_id: str) -> None:
    st.session_state["job_id"] = job_id


def _current_job() -> str | None:
    val = st.session_state.get("job_id")
    return str(val) if val else None


def render() -> None:
    """トップページの描画."""
    st.set_page_config(page_title="Excel改修支援ツール", layout="wide")
    st.title("📊 Excel改修支援ツール")
    st.caption(
        "VBA/数式/参照関係を含む `.xlsm` `.xls` ツールの統合設計書を生成し、改修対話を支援します。"
    )

    client = _client()

    # ---------- バックエンド接続確認 ----------
    if not client.health():
        st.error(
            f"Backend ({client.base_url}) に接続できません。"
            " `uv run uvicorn backend.main:app --reload --port 8000` を起動してください。"
        )
        return

    # ---------- アップロード ----------
    st.subheader("ファイルアップロード")
    uploaded = st.file_uploader(
        "改修対象の Excel ファイルを選択",
        type=["xlsm", "xls", "xlsx"],
        accept_multiple_files=False,
    )

    if uploaded is not None and st.button("アップロード & 抽出を開始", type="primary"):
        try:
            with st.spinner("抽出中..."):
                job_id = client.extract(uploaded.name, uploaded.getvalue())
            _set_current_job(job_id)
            st.success(f"ジョブを作成しました: {job_id}")
            st.info("左のサイドバーから Upload / Spec / Chat ページに進んでください。")
        except BackendError as e:
            st.error(f"抽出に失敗しました: {e}")

    st.divider()

    # ---------- 過去ジョブ一覧 ----------
    st.subheader("過去のジョブ")
    try:
        jobs = client.list_jobs()
    except BackendError as e:
        st.error(f"ジョブ一覧の取得に失敗: {e}")
        return

    if not jobs:
        st.info("まだジョブがありません。上のフォームからファイルをアップロードしてください。")
        return

    current = _current_job()
    for job in jobs:
        cols = st.columns([4, 2, 2, 1, 1])
        with cols[0]:
            mark = "✅" if current == job["job_id"] else ""
            st.write(f"{mark} **{job['filename']}**  \n`{job['job_id']}`")
        with cols[1]:
            st.write(f"作成: {job['created_at']}")
        with cols[2]:
            st.write(f"状態: `{job['status']}`")
        with cols[3]:
            if st.button("選択", key=f"select_{job['job_id']}"):
                _set_current_job(job["job_id"])
                st.rerun()
        with cols[4]:
            if st.button("削除", key=f"delete_{job['job_id']}"):
                try:
                    client.delete_job(job["job_id"])
                    if current == job["job_id"]:
                        st.session_state.pop("job_id", None)
                    st.rerun()
                except BackendError as e:
                    st.error(f"削除失敗: {e}")


if __name__ == "__main__" or True:  # Streamlit は import 経由で実行される
    render()
