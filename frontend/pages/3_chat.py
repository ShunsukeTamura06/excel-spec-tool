"""改修対話 (チャット) ページ."""

from __future__ import annotations

import streamlit as st

from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def render() -> None:
    st.set_page_config(page_title="Chat", layout="wide")
    st.title("💬 Chat — 改修対話")

    job_id = st.session_state.get("job_id")
    if not job_id:
        st.warning("ジョブが選択されていません。トップで選択してください。")
        return

    client = _client()

    # 履歴ロード
    try:
        history = client.get_chat_history(job_id)
    except BackendError as e:
        if e.status_code == 404:
            st.error("ジョブが見つかりません。")
        else:
            st.error(f"履歴取得失敗: {e}")
        return

    st.caption("ジョブ ID: `" + job_id + "`")

    # 履歴表示
    for msg in history:
        role = msg.get("role", "user")
        content = msg.get("content", "")
        with st.chat_message(role):
            st.markdown(content)

    # 入力
    user_input = st.chat_input("改修したい内容や質問を入力 (例: H2 セルの計算式を変えたい)")
    if user_input:
        with st.chat_message("user"):
            st.markdown(user_input)
        try:
            with st.spinner("LLM 応答中..."):
                result = client.chat(job_id, user_input)
        except BackendError as e:
            st.error(f"チャット失敗: {e}")
            return
        with st.chat_message("assistant"):
            st.markdown(result.get("reply", ""))
        # 次のレンダで履歴が再取得される
        st.rerun()


render()
