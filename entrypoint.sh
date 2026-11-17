#!/bin/bash
set -e

echo "Starting Barcode Tracker..."

# Run dws auth login with service account credentials if env vars are set
if [ -n "$DINGTALK_CLIENT_ID" ] && [ -n "$DINGTALK_CLIENT_SECRET" ]; then
    echo "Authenticating dws-core with service account..."
    dws auth login --client-id "$DINGTALK_CLIENT_ID" --client-secret "$DINGTALK_CLIENT_SECRET" || {
        echo "Warning: dws auth login failed, continuing anyway..."
    }
fi

# Start gunicorn
exec gunicorn app:app --bind 0.0.0.0:8000 --workers 2 --timeout 120
