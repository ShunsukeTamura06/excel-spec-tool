"""設計書ページ — 統合設計書 Markdown 表示 + 参照逆引き検索."""

from __future__ import annotations

import streamlit as st

from frontend._sidebar import render_job_sidebar
from frontend.api_client import BackendClient, BackendError


def _client() -> BackendClient:
    if "_backend_client" not in st.session_state:
        st.session_state["_backend_client"] = BackendClient()
    client: BackendClient = st.session_state["_backend_client"]
    return client


st.title("📑 設計書")

job_id = render_job_sidebar()
if not job_id:
    st.warning("ジョブが選択されていません。")
    if st.button("🏠 ホームへ"):
        st.switch_page("views/home.py")
    st.stop()

client = _client()

try:
    body = client.get_spec(job_id)
except BackendError as e:
    if e.status_code == 409:
        st.warning("設計書がまだ生成されていません。ホームから再分析してください。")
    elif e.status_code == 404:
        st.error("ジョブが見つかりません。")
    else:
        st.error(f"取得失敗: {e}")
    st.stop()

meta = body.get("meta", {})
st.caption(
    f"**{meta.get('filename', '?')}** / "
    f"状態: `{meta.get('status', '?')}` / "
    f"作成: {meta.get('created_at', '?')[:19]}"
)

spec_md = body.get("spec_md", "")
st.divider()
st.markdown(spec_md, unsafe_allow_html=True)

st.divider()
st.subheader("🔍 参照逆引き検索")
st.caption("セルや範囲 (例: `Calc!H2`, `Input!A:A`) を参照している箇所を検索します。")

target = st.text_input("参照先キー", placeholder="例: Calc!H2", key="ref_target")
if target and st.button("検索", key="ref_search"):
    try:
        refs = client.get_references(job_id, target)
    except BackendError as e:
        st.error(f"検索失敗: {e}")
    else:
        if not refs:
            st.info(f"`{target}` を参照している箇所は見つかりませんでした。")
        else:
            st.success(f"`{target}` を参照: {len(refs)} 件")
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
                hide_index=True,
            )
