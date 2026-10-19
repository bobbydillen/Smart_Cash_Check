FROM python:3.11-slim
WORKDIR /app
COPY requirements-streamlit.txt .
RUN pip install --no-cache-dir -r requirements-streamlit.txt
COPY . .
ENV HOSTED=1
EXPOSE 8501
# Render/Railway/Fly inject $PORT; default to 8501 elsewhere
CMD streamlit run streamlit_app.py --server.port=${PORT:-8501} --server.address=0.0.0.0 --server.headless=true --server.enableXsrfProtection=true
