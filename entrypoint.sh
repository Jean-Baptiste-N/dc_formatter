#!/bin/bash
set -e

# Fix permissions for bind-mounted volumes at runtime
# This allows dcformatter user to write to volumes created by host user
for dir in DC_SOURCES OUTPUTS_FORMATTED; do
    if [ -d "$dir" ]; then
        echo "Fixing permissions for $dir..."
        chmod 777 "$dir" 2>/dev/null || true
        chown -R dcformatter:dcformatter "$dir" 2>/dev/null || true
    fi
done

# Ensure ephemeral volume directories have proper ownership
for dir in OUTPUT1_XML-RAW OUTPUT2_JSON-RAW OUTPUT3_JSON-TRANSFORMED OUTPUT4_DOCX-RESULT TEMPLATE; do
    if [ -d "$dir" ]; then
        chown -R dcformatter:dcformatter "$dir" 2>/dev/null || true
        chmod 777 "$dir" 2>/dev/null || true
    fi
done

echo "✓ Permissions fixed, starting application..."

# Run as dcformatter user using gosu
exec gosu dcformatter python -m uvicorn app:app --host 0.0.0.0 --port 8000 --reload
