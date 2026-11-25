# Python plus the exact R package families used by the uploaded Venn/Manhattan/QBA scripts.
FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

RUN apt-get update && apt-get install -y --no-install-recommends \
    r-base-core \
    r-cran-readxl r-cran-dplyr r-cran-tidyr r-cran-venndiagram \
    r-cran-openxlsx r-cran-tibble r-cran-ggplot2 r-cran-stringr \
    r-cran-ggrepel r-cran-gridextra r-cran-lubridate r-cran-purrr \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /opt/mitchell-web
COPY requirements-web.txt ./
RUN pip install --no-cache-dir -r requirements-web.txt
COPY . .
RUN useradd -r -m -u 10001 mitchell && chown -R mitchell:mitchell /opt/mitchell-web
USER mitchell
EXPOSE 8501
CMD ["sh", "-c", "exec streamlit run streamlit_app.py --server.address=0.0.0.0 --server.port=${PORT:-8501} --server.headless=true"]
