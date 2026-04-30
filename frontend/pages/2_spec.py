"""設計書 (Markdown) 表示 + 参照逆引き検索ページ."""

from __future__ import annotations

import streamlit as st

from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


def render() -> None:
    st.set_page_config(page_title="Spec", layout="wide")
    st.title("📑 Spec — 統合設計書")

    job_id = st.session_state.get("job_id")
    if not job_id:
        st.warning("ジョブが選択されていません。トップで選択してください。")
        return

    client = _client()

    try:
        body = client.get_spec(job_id)
    except BackendError as e:
        if e.status_code == 409:
            st.info("設計書がまだ生成されていません。Upload ページで解析を実行してください。")
        elif e.status_code == 404:
            st.error("ジョブが見つかりません。")
        else:
            st.error(f"設計書取得に失敗: {e}")
        return

    meta = body.get("meta", {})
    st.caption(
        f"**{meta.get('filename', '?')}** / 状態: `{meta.get('status', '?')}` "
        f"/ 作成: {meta.get('created_at', '?')}"
    )

    spec_md = body.get("spec_md", "")
    st.divider()

    # 設計書本文
    with st.container():
        st.markdown(spec_md, unsafe_allow_html=True)

    st.divider()
    st.subheader("🔍 参照逆引き検索")
    st.caption("セルや範囲 (例: `Calc!H2`, `Input!A:A`) を参照している箇所を検索します。")

    target = st.text_input("参照先キー", placeholder="例: Calc!H2")
    if target and st.button("検索"):
        try:
            refs = client.get_references(job_id, target)
        except BackendError as e:
            st.error(f"検索失敗: {e}")
            return
        if not refs:
            st.info(f"`{target}` を参照している箇所は見つかりませんでした。")
        else:
            st.write(f"`{target}` を参照している箇所: {len(refs)} 件")
            st.dataframe(
                [
                    {
                        "kind": r.get("kind"),
                        "from": r.get("from"),
                        "code": r.get("code", ""),
                    }
                    for r in refs
                ],
                use_container_width=True,
            )


render()
