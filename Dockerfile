FROM mcr.microsoft.com/playwright:v1.54.0-noble
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
ENV NODE_OPTIONS="--max-old-space-size=8192"
CMD ["npm", "start"]