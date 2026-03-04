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

echo "🔑 API Key loaded (length: ${#QBTOJSON_API_KEY})"

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
  --timeout 300s \
  --max-instances 50 \
  --min-instances 0 \
  --set-env-vars QBTOJSON_API_KEY="${QBTOJSON_API_KEY}"

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
