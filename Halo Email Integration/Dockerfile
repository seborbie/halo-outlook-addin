FROM oven/bun:1.3.14-alpine AS web-build

WORKDIR /workspace
COPY package.json bun.lock ./
COPY apps/web/package.json apps/web/package.json
RUN bun install --frozen-lockfile
COPY apps/web apps/web
RUN bun run build:web

FROM node:22-bookworm-slim AS addin-build

WORKDIR /workspace/Halo Email Integration
COPY ["Halo Email Integration/package.json", "Halo Email Integration/package-lock.json", "./"]
RUN npm ci
COPY ["Halo Email Integration", "./"]

ARG PUBLIC_BASE_URL
ENV PUBLIC_BASE_URL=${PUBLIC_BASE_URL}
RUN npm run build

FROM node:22-bookworm-slim AS runtime

ENV NODE_ENV=production
ENV PORT=3000
WORKDIR /app

COPY ["Halo Email Integration/package.json", "Halo Email Integration/package-lock.json", "./"]
RUN npm ci --omit=dev
COPY --from=addin-build ["/workspace/Halo Email Integration/server", "./server"]
COPY --from=addin-build ["/workspace/Halo Email Integration/dist", "./dist"]
COPY --from=web-build /workspace/apps/web/build ./web

EXPOSE 3000
CMD ["npm", "run", "serve"]
