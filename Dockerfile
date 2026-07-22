# Container image for hosting anywhere that runs Docker (Railway, Fly.io, Cloud Run, etc.)
FROM node:20-alpine
WORKDIR /app

# Install dependencies first (better layer caching)
COPY package*.json ./
RUN npm install --omit=dev

# App source
COPY . .

# The app reads PORT from the environment (defaults to 3000)
ENV PORT=3000
EXPOSE 3000

CMD ["npm", "start"]
