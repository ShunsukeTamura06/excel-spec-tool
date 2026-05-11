"""Streamlit エントリ. ナビゲーション定義のみ.

ページ本体は `frontend/views/` 配下に分離している。

起動例:
    uv run streamlit run frontend/app.py --server.port 8501
"""

from __future__ import annotations

import streamlit as st

st.set_page_config(
    page_title="Excel 改修支援ツール",
    layout="wide",
)

home_page = st.Page(
    "views/home.py",
    title="ホーム",
    icon="🏠",
    default=True,
)
spec_page = st.Page(
    "views/spec.py",
    title="設計書",
    icon="📑",
)
chat_page = st.Page(
    "views/chat.py",
    title="チャット",
    icon="💬",
)

pg = st.navigation([home_page, spec_page, chat_page])
pg.run()
