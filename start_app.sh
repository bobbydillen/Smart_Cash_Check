#!/bin/sh
cd "$(dirname "$0")"
python3 -m pip install -r requirements-streamlit.txt
python3 -m streamlit run streamlit_app.py
