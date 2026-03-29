FROM rocker/shiny:4.4.1

# Install system dependencies for R packages
RUN apt-get update && apt-get install -y \
    libxml2-dev \
    libcurl4-openssl-dev \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Install R packages
RUN R -e "install.packages(c('shiny', 'bslib', 'DT', 'tidyxl', 'openxlsx2', 'readxl'), repos='https://cloud.r-project.org/')"

# Copy app files
COPY app.R /srv/shiny-server/app.R
COPY R/ /srv/shiny-server/R/
COPY inst/ /srv/shiny-server/inst/

# Cloud Run uses PORT env variable (default 8080)
ENV PORT=8080

# Run Shiny on the port Cloud Run expects
CMD ["R", "-e", "shiny::runApp('/srv/shiny-server', host='0.0.0.0', port=as.numeric(Sys.getenv('PORT', 8080)))"]
