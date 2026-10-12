FROM node:22.16.0-bookworm-slim@sha256:048ed02c5fd52e86fda6fbd2f6a76cf0d4492fd6c6fee9e2c463ed5108da0e34
COPY resources/debian.sources /etc/apt/sources.list.d/debian.sources
RUN apt-get update -qq && DEBIAN_FRONTEND=noninteractive apt-get install -y -qq --no-install-recommends \
    texlive-xetex texlive-latex-extra texlive-fonts-recommended \
    poppler-utils libreoffice-writer fonts-liberation \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /opt/proceedings-md
COPY package.json package-lock.json ./
RUN npm ci --ignore-scripts
COPY tools tools
COPY resources resources
WORKDIR /work
ENTRYPOINT ["/opt/proceedings-md/node_modules/.bin/tsx", "/opt/proceedings-md/tools/cli.ts"]
CMD ["/work/sample/sample.md", "/work/build/sample.docx"]
