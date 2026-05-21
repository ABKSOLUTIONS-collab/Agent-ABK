# =============================================================================
# Agent 365 Bridge — Multi-Stage Docker Build
# Stage 1: Builder (TypeScript compile)
# Stage 2: Runtime (lean Node.js image, no dev deps)
# =============================================================================

# ── Stage 1: Builder ──────────────────────────────────────────────────────────
FROM node:22-alpine AS builder

WORKDIR /app

# Copy package files first (layer caching)
COPY package.json package-lock.json ./

# Install ALL deps (including devDependencies for TypeScript compile)
RUN npm ci

# Copy source code
COPY tsconfig.json ./
COPY src/ ./src/
COPY ToolingManifest.json ./

# Compile TypeScript → dist/
RUN npm run build


# ── Stage 2: Runtime ─────────────────────────────────────────────────────────
FROM node:22-alpine AS runtime

WORKDIR /app

# Copy package files
COPY package.json package-lock.json ./

# Install PRODUCTION deps only (no TypeScript, no ts-node)
RUN npm ci --omit=dev

# Copy compiled output from builder
COPY --from=builder /app/dist ./dist

# Copy manifest (read at runtime by configuration.ts)
COPY ToolingManifest.json ./

# ── Auth cache directory ──────────────────────────────────────────────────────
# In Azure Container Apps we use client_credentials (no interactive login).
# The ~/.agent365-bridge dir is kept for optional volume-mount scenarios.
RUN mkdir -p /root/.agent365-bridge

# ── Runtime config ────────────────────────────────────────────────────────────
ENV NODE_ENV=production

# MCP servers communicate over stdio — no port needed.
# Azure Container Apps will inject all env vars as secrets.

# Healthcheck: HTTP GET /health — uses Node's built-in fetch (Node 22+)
HEALTHCHECK --interval=30s --timeout=10s --start-period=20s --retries=3 \
  CMD node -e "fetch('http://localhost:3000/health').then(r=>r.ok?process.exit(0):process.exit(1)).catch(()=>process.exit(1))"

CMD ["node", "dist/index.js"]