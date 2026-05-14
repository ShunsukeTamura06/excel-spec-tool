"""共通サイドバー部品.

各ページのサイドバー上部に「現在のジョブ」インジケータと
「ホームに戻る」ボタンを表示する。
"""

from __future__ import annotations

import streamlit as st

from frontend_streamlit.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def render_job_sidebar(show_home_button: bool = True) -> str | None:
    """サイドバー上部にカレントジョブの情報を出す.

    Args:
        show_home_button: 「ホームに戻る」ボタンを表示するか. ホームページでは False.

    Returns:
        現在の job_id. 未選択 / 不整合の場合は None.
    """
    job_id = st.session_state.get("job_id")

    with st.sidebar:
        st.divider()
        if not job_id:
            st.warning("⚠️ ジョブ未選択")
            st.caption("🏠 ホームから Excel ファイルを開いてください")
            return None

        try:
            jobs = _client().list_jobs()
            current = next((j for j in jobs if j["job_id"] == job_id), None)
        except BackendError:
            st.error("⚠️ Backend 接続エラー")
            return job_id

        if not current:
            st.error("⚠️ 選択中のジョブが見つかりません")
            st.session_state.pop("job_id", None)
            return None

        st.success(f"📂 **{current['filename']}**")
        st.caption(f"状態: `{current['status']}`")
        st.caption(f"`{job_id[:8]}...`")

        if show_home_button:
            if st.button("↩️ ホームに戻る", use_container_width=True):
                st.switch_page("views/home.py")

    return job_id
