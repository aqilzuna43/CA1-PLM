#!/bin/bash
# Backup current production code

BACKUP_DIR="backups/$(date '+%Y%m%d_%H%M%S')"
mkdir -p "$BACKUP_DIR"

echo "💾 Creating backup in $BACKUP_DIR..."

# Pull current code from Google
clasp pull

# Copy to backup directory
cp -r src "$BACKUP_DIR/"
cp appsscript.json "$BACKUP_DIR/"

echo "✅ Backup created successfully!"
echo "📁 Location: $BACKUP_DIR"
