FROM oven/bun:1.3 AS build
WORKDIR /app
COPY package.json bun.lock ./
RUN bun install --frozen-lockfile
COPY . .
RUN bun run build

FROM oven/bun:1.3-slim AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV PORT=9812
ENV HOST=0.0.0.0
COPY --from=build /app/.output ./.output
COPY --from=build /app/package.json ./package.json
EXPOSE 9812
CMD ["bun", "run", ".output/server/index.mjs"]
