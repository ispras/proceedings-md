FROM node:23-slim

WORKDIR /app
COPY . .

RUN apt-get update && \
    apt-get install -y pandoc && \
    npm install --production && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

ENTRYPOINT ["./run-local.sh"]
