#!/bin/bash
# Deploy script for CA1-PLM

echo "🚀 Deploying CA1-PLM to Google Apps Script..."

# Check if clasp is logged in
if ! clasp login --status 2>/dev/null; then
    echo "❌ Not logged in to clasp. Run: npm run login"
    exit 1
fi

# Pull latest from Google (backup)
echo "📥 Pulling latest from Google Apps Script..."
clasp pull

# Push local changes
echo "📤 Pushing changes to Google Apps Script..."
clasp push

# Create a new version
echo "📦 Creating new version..."
VERSION=$(clasp version "Auto-deploy $(date '+%Y-%m-%d %H:%M:%S')" | grep -oE '[0-9]+$')

echo "✅ Deployed successfully! Version: $VERSION"
echo "🌐 Open in browser: npm run open"
