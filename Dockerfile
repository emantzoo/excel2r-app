FROM rocker/shiny:4.4.1

# Install system dependencies for R packages
RUN apt-get update && apt-get install -y \
    libxml2-dev \
    libcurl4-openssl-dev \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Install R package dependencies
RUN R -e "install.packages(c('shiny', 'bslib', 'DT', 'tidyxl', 'openxlsx2', 'readxl'), repos='https://cloud.r-project.org/')"

# Copy package source and install it
COPY . /tmp/excel2r
RUN R CMD INSTALL /tmp/excel2r && rm -rf /tmp/excel2r

# Cloud Run uses PORT env variable (default 8080)
ENV PORT=8080

# Run the Shiny app via the package
CMD ["R", "-e", "excel2r::run_app(host='0.0.0.0', port=as.numeric(Sys.getenv('PORT', 8080)))"]
