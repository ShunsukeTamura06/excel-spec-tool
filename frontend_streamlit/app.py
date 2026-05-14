"""Streamlit エントリ (LEGACY).

Nuxt 移行中. このコードは参考用に一時保存. パリティ到達後に削除予定.

起動例:
    uv run streamlit run frontend_streamlit/app.py --server.port 8501
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
