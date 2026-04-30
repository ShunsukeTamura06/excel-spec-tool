"""アップロード→抽出→解析の進捗表示ページ.

トップで作成 (or 選択) したジョブに対して、解析 (`/analyze`) を実行する。
"""

from __future__ import annotations

import streamlit as st

from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def render() -> None:
    st.title("⬆️ Upload — 抽出と解析")

    job_id = st.session_state.get("job_id")
    if not job_id:
        st.warning("ジョブが選択されていません。トップページでアップロードまたは選択してください。")
        return

    client = _client()

    try:
        jobs = client.list_jobs()
    except BackendError as e:
        st.error(f"ジョブ取得に失敗: {e}")
        return

    job = next((j for j in jobs if j["job_id"] == job_id), None)
    if job is None:
        st.error("選択中のジョブがバックエンドに存在しません。トップで再選択してください。")
        return

    st.write(f"**ファイル**: {job['filename']}")
    st.write(f"**ジョブID**: `{job_id}`")
    st.write(f"**現在の状態**: `{job['status']}`")
    st.divider()

    if job["status"] == "uploaded":
        st.info("抽出が完了していません。再アップロードしてください。")
    elif job["status"] == "failed":
        st.error("抽出または解析が失敗しています。ログを確認してください。")
    elif job["status"] == "extracted":
        st.info("Core 抽出が完了しました。次は LLM 注釈 + 設計書生成です。")
        if st.button("解析を実行 (POST /analyze)", type="primary"):
            try:
                with st.spinner("解析中..."):
                    client.analyze(job_id)
                st.success("解析が完了しました。Spec ページで設計書を確認できます。")
                st.rerun()
            except BackendError as e:
                st.error(f"解析失敗: {e}")
    elif job["status"] == "analyzed":
        st.success("解析済みです。Spec / Chat ページに進んでください。")
        if st.button("再解析"):
            try:
                with st.spinner("再解析中..."):
                    client.analyze(job_id)
                st.rerun()
            except BackendError as e:
                st.error(f"解析失敗: {e}")


render()
