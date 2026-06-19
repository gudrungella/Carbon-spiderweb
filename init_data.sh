#!/bin/bash
# Copies templates into data/ on first use — skips files that already exist.
mkdir -p data
for f in templates/*.xlsx; do
    target="data/$(basename "$f")"
    if [ -f "$target" ]; then
        echo "Skipping $target (already exists)"
    else
        cp "$f" "$target"
        echo "Created $target"
    fi
done
