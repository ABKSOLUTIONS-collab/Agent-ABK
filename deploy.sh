#!/usr/bin/env bash
# =============================================================================
# Agent 365 Bridge — Azure Deployment Script
# =============================================================================

set -euo pipefail

# ── Configuration ─────────────────────────────────────────────────────────────
RESOURCE_GROUP="ABKAgent365"
LOCATION="swedencentral"
ACR_NAME="Agent365Registry"
ACA_ENV_NAME="env-agent365-bridge"
APP_NAME="agent365-bridge"
IMAGE_TAG=$(date +%Y%m%d%H%M%S)
STORAGE_ACCOUNT_NAME="abkagent365storage"
STORAGE_SHARE_NAME="agent365-tokens"
VOLUME_NAME="token-storage"
VOLUME_MOUNT_PATH="/root/.agent365-bridge"

# ── Server URL (hardcoded — update here if Azure reassigns the domain) ────────
SERVER_BASE_URL="https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io"

# ── Secrets (must be exported in the calling shell before running) ────────────
: "${AZURE_TENANT_ID:?Need AZURE_TENANT_ID}"
: "${AZURE_CLIENT_ID:?Need AZURE_CLIENT_ID}"
: "${AZURE_CLIENT_SECRET:?Need AZURE_CLIENT_SECRET}"
: "${MCP_API_KEY:?Need MCP_API_KEY}"
: "${AZURE_DI_ENDPOINT:?Need AZURE_DI_ENDPOINT}"
: "${AZURE_DI_KEY:?Need AZURE_DI_KEY}"

# ── Derived ───────────────────────────────────────────────────────────────────
ACR_SERVER="${ACR_NAME,,}.azurecr.io"
IMAGE_FULL="${ACR_SERVER}/${APP_NAME}:${IMAGE_TAG}"

echo "🔧  Resource Group  : $RESOURCE_GROUP"
echo "🏗   ACR              : $ACR_NAME"
echo "📦  Image             : $IMAGE_FULL"
echo "💾  Storage Account  : $STORAGE_ACCOUNT_NAME"
echo ""

# ── 1. ACR credentials ────────────────────────────────────────────────────────
echo "▶ Getting ACR credentials..."
ACR_USERNAME=$(az acr credential show --name "$ACR_NAME" --query username -o tsv)
ACR_PASSWORD=$(az acr credential show --name "$ACR_NAME" --query passwords[0].value -o tsv)

# ── 2. Docker Build ───────────────────────────────────────────────────────────
echo "▶ Building Docker image..."
docker buildx build --platform linux/amd64 --provenance=false --no-cache -t "$IMAGE_FULL" .

# ── 3. Login to ACR and Push ──────────────────────────────────────────────────
echo "▶ Logging in to ACR..."
az acr login --name "$ACR_NAME"

echo "▶ Pushing image to ACR..."
docker push "$IMAGE_FULL"

# ── 4. Storage Account ────────────────────────────────────────────────────────
echo "▶ Creating Storage Account for token persistence..."
az storage account create \
  --name "$STORAGE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Standard_LRS \
  --output none 2>/dev/null || echo "   Storage account already exists, skipping."

STORAGE_KEY=$(az storage account keys list \
  --account-name "$STORAGE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query "[0].value" -o tsv)

echo "▶ Creating file share..."
az storage share create \
  --name "$STORAGE_SHARE_NAME" \
  --account-name "$STORAGE_ACCOUNT_NAME" \
  --account-key "$STORAGE_KEY" \
  --output none 2>/dev/null || echo "   File share already exists, skipping."

# ── 5. ACA Environment ────────────────────────────────────────────────────────
echo "▶ Checking ACA Environment..."
if az containerapp env show --name "$ACA_ENV_NAME" --resource-group "$RESOURCE_GROUP" --output none 2>/dev/null; then
  echo "   ACA environment already exists, skipping."
else
  echo "   Creating ACA environment..."
  az containerapp env create \
    --name "$ACA_ENV_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --location "$LOCATION" \
    --output none
fi

# ── 6. Link Storage to ACA Environment ───────────────────────────────────────
echo "▶ Linking storage to ACA environment..."
az containerapp env storage set \
  --name "$ACA_ENV_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --storage-name "$VOLUME_NAME" \
  --azure-file-account-name "$STORAGE_ACCOUNT_NAME" \
  --azure-file-account-key "$STORAGE_KEY" \
  --azure-file-share-name "$STORAGE_SHARE_NAME" \
  --access-mode ReadWrite \
  --output none

# ── 7. Container App ──────────────────────────────────────────────────────────
echo "▶ Deploying Container App..."

if az containerapp show --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" &>/dev/null; then
  echo "   Updating existing Container App..."
  # Set registry credentials separately (az containerapp update doesn't support --registry-* flags)
  az containerapp registry set \
    --name "$APP_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --server "$ACR_SERVER" \
    --username "$ACR_USERNAME" \
    --password "$ACR_PASSWORD" \
    --output none
  az containerapp update \
    --name "$APP_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --image "$IMAGE_FULL" \
    --set-env-vars \
        NODE_ENV=production \
        AZURE_TENANT_ID="$AZURE_TENANT_ID" \
        AZURE_CLIENT_ID="$AZURE_CLIENT_ID" \
        AZURE_CLIENT_SECRET="${AZURE_CLIENT_SECRET}" \
        MCP_PLATFORM_ENDPOINT="https://agent365.svc.cloud.microsoft" \
        MCP_PLATFORM_AUTHENTICATION_SCOPE="ea9ffc3e-8a23-4a7d-836d-234d7c7565c1/.default" \
        MCP_API_KEY="$MCP_API_KEY" \
        SERVER_BASE_URL="$SERVER_BASE_URL" \
        AZURE_STORAGE_ACCOUNT="$STORAGE_ACCOUNT_NAME" \
        AZURE_STORAGE_KEY="$STORAGE_KEY" \
        TOKEN_TABLE_NAME="agent365tokens" \
        AZURE_DI_ENDPOINT="$AZURE_DI_ENDPOINT" \
        AZURE_DI_KEY="${AZURE_DI_KEY}"
else
  echo "   Creating new Container App..."
  az containerapp create \
    --name "$APP_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --environment "$ACA_ENV_NAME" \
    --image "$IMAGE_FULL" \
    --registry-server "$ACR_SERVER" \
    --registry-username "$ACR_USERNAME" \
    --registry-password "$ACR_PASSWORD" \
    --target-port 3000 \
    --ingress external \
    --min-replicas 1 \
    --max-replicas 1 \
    --cpu 0.25 \
    --memory 0.5Gi \
    --env-vars \
        NODE_ENV=production \
        AZURE_TENANT_ID="$AZURE_TENANT_ID" \
        AZURE_CLIENT_ID="$AZURE_CLIENT_ID" \
        AZURE_CLIENT_SECRET="${AZURE_CLIENT_SECRET}" \
        MCP_PLATFORM_ENDPOINT="https://agent365.svc.cloud.microsoft" \
        MCP_PLATFORM_AUTHENTICATION_SCOPE="ea9ffc3e-8a23-4a7d-836d-234d7c7565c1/.default" \
        MCP_API_KEY="$MCP_API_KEY" \
        SERVER_BASE_URL="$SERVER_BASE_URL" \
        AZURE_STORAGE_ACCOUNT="$STORAGE_ACCOUNT_NAME" \
        AZURE_STORAGE_KEY="$STORAGE_KEY" \
        TOKEN_TABLE_NAME="agent365tokens" \
        AZURE_DI_ENDPOINT="$AZURE_DI_ENDPOINT" \
        AZURE_DI_KEY="${AZURE_DI_KEY}"
fi

# ── 8. Get URL ────────────────────────────────────────────────────────────────
FQDN=$(az containerapp show \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query properties.configuration.ingress.fqdn -o tsv)

echo ""
echo "✅  Deployment complete! (image tag: $IMAGE_TAG)"
echo ""
echo "🌐  MCP URL: https://${FQDN}/mcp"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "👉  NEXT STEP: Get your personal URL"
echo ""
echo "   Open this in your browser and sign in:"
echo "   https://${FQDN}/login"
echo ""
echo "   You will get your permanent personal URL to add to Claude.ai"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"