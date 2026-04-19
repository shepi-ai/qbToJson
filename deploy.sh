#!/bin/bash
# Deployment script for qbToJson to Google Cloud Run
# Uses Cloud Build (--source) to avoid registry issues

set -e  # Exit on error

# Configuration
SERVICE_NAME="qbtojson"
REGION="us-central1"
GCLOUD="/opt/homebrew/share/google-cloud-sdk/bin/gcloud"

echo "🚀 Deploying qbToJson to Google Cloud Run"
echo "=========================================="
echo "Service: ${SERVICE_NAME}"
echo "Region: ${REGION}"
echo "Method: Cloud Build (--source)"
echo ""

# Check if gcloud is installed
if [ ! -f "$GCLOUD" ]; then
    echo "❌ Error: gcloud CLI not found at $GCLOUD"
    echo "Please update GCLOUD variable in deploy.sh with the correct path"
    exit 1
fi

# Check if logged in to gcloud
echo "📋 Checking gcloud authentication..."
if ! $GCLOUD auth list --filter=status:ACTIVE --format="value(account)" | grep -q .; then
    echo "❌ Error: Not authenticated with gcloud. Run: $GCLOUD auth login"
    exit 1
fi

# Load API key from .env file
if [ -f .env ]; then
    export $(grep -v '^#' .env | xargs)
    echo "✅ Loaded environment variables from .env"
else
    echo "⚠️  Warning: .env file not found"
fi

# Check if QBTOJSON_API_KEY is set
if [ -z "$QBTOJSON_API_KEY" ]; then
    echo "❌ Error: QBTOJSON_API_KEY not set in .env file"
    exit 1
fi

# Load Anthropic API key from local .env or fallback to llmextractorapi .env
if [ -z "$ANTHROPIC_API_KEY" ]; then
    LLM_ENV_PATH="/Users/araboin/shepi/llmextractorapi/.env"
    if [ -f "$LLM_ENV_PATH" ]; then
        ANTHROPIC_API_KEY=$(grep '^ANTHROPIC_API_KEY=' "$LLM_ENV_PATH" | cut -d= -f2-)
        export ANTHROPIC_API_KEY
        if [ -n "$ANTHROPIC_API_KEY" ]; then
            echo "✅ Loaded ANTHROPIC_API_KEY from $LLM_ENV_PATH"
        fi
    fi
fi

# Defaults for GL PDF LLM fallback
ENABLE_LLM_GL_PDF_FALLBACK=${ENABLE_LLM_GL_PDF_FALLBACK:-true}
GL_PDF_FORCE_LLM=${GL_PDF_FORCE_LLM:-true}
GL_LLM_MODEL=${GL_LLM_MODEL:-claude-haiku-4-5-20251001}
GL_LLM_MAX_ATTEMPTS=${GL_LLM_MAX_ATTEMPTS:-5}

# Check if ANTHROPIC_API_KEY is set (required for PDF fallback)
if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "⚠️  Warning: ANTHROPIC_API_KEY not found. GL PDF LLM fallback will be disabled."
    ENABLE_LLM_GL_PDF_FALLBACK=false
fi

echo "🔑 API Key loaded (length: ${#QBTOJSON_API_KEY})"
if [ -n "$ANTHROPIC_API_KEY" ]; then
echo "🤖 Anthropic key loaded (length: ${#ANTHROPIC_API_KEY})"
fi
echo "🤖 GL PDF fallback enabled: ${ENABLE_LLM_GL_PDF_FALLBACK}"
echo "🤖 GL PDF force LLM: ${GL_PDF_FORCE_LLM}"
echo "🤖 GL model: ${GL_LLM_MODEL}"
echo "🤖 GL max attempts: ${GL_LLM_MAX_ATTEMPTS}"

# Deploy to Cloud Run using Cloud Build
echo ""
echo "🚀 Deploying to Cloud Run with Cloud Build..."
echo "   (This will build and deploy automatically)"
echo ""

$GCLOUD run deploy ${SERVICE_NAME} \
  --source . \
  --region ${REGION} \
  --allow-unauthenticated \
  --memory 2Gi \
  --cpu 2 \
  --concurrency 1 \
  --cpu-boost \
  --no-cpu-throttling \
  --timeout 540s \
  --max-instances 50 \
  --min-instances 2 \
  --set-env-vars QBTOJSON_API_KEY="${QBTOJSON_API_KEY}",SUPABASE_URL="${SUPABASE_URL}",SUPABASE_SERVICE_ROLE_KEY="${SUPABASE_SERVICE_ROLE_KEY}",ENABLE_LLM_GL_PDF_FALLBACK="${ENABLE_LLM_GL_PDF_FALLBACK}",GL_PDF_FORCE_LLM="${GL_PDF_FORCE_LLM}",ANTHROPIC_API_KEY="${ANTHROPIC_API_KEY}",GL_LLM_MODEL="${GL_LLM_MODEL}",GL_LLM_MAX_ATTEMPTS="${GL_LLM_MAX_ATTEMPTS}"

if [ $? -ne 0 ]; then
    echo "❌ Cloud Run deployment failed!"
    exit 1
fi

# Get the service URL
echo ""
echo "✅ Deployment successful!"
echo ""
SERVICE_URL=$($GCLOUD run services describe ${SERVICE_NAME} --region ${REGION} --format 'value(status.url)')
echo "🌐 Service URL: ${SERVICE_URL}"
echo ""
echo "📋 Test the service:"
echo "   Health check: curl ${SERVICE_URL}/health"
echo "   API info: curl ${SERVICE_URL}/api/info"
echo ""
echo "🎉 Deployment complete!"
