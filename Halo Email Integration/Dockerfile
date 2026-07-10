FROM node:22-bookworm-slim

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .

ARG PUBLIC_BASE_URL
ENV PUBLIC_BASE_URL=${PUBLIC_BASE_URL}
RUN npm run build

ENV NODE_ENV=production
ENV PORT=3000

EXPOSE 3000

CMD ["npm", "run", "serve"]
