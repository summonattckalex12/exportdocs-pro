FROM oven/bun:1.3 AS build
WORKDIR /app
COPY package.json bun.lock ./
RUN bun install --frozen-lockfile
COPY . .
RUN bun run build

FROM node:22-slim AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV PORT=9812
ENV HOST=0.0.0.0
# nitro node-server preset emits a self-contained server under .output/
COPY --from=build /app/.output ./.output
COPY --from=build /app/package.json ./package.json
RUN mkdir -p /app/data/exports /app/data/uploads
EXPOSE 9812
CMD ["node", ".output/server/index.mjs"]
