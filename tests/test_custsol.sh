#!/usr/bin/env bash
# Compile the two Cust Sol Office Scripts to plain JS and run the offline
# test harness against them. Requires node + a local typescript install.
set -u
cd "$(dirname "$0")/.."

BUILD_DIR="$(mktemp -d)"
trap 'rm -rf "$BUILD_DIR"' EXIT

# Compiled separately: both files declare their own main(), which tsc treats
# as duplicate declarations when compiled together. The ExcelScript namespace
# errors are expected (no local type declarations) — emission still succeeds.
npx --no-install tsc --skipLibCheck --target es2020 --lib es2020,dom \
  --outDir "$BUILD_DIR/reader" office_scripts/custsol_read_daily.ts >/dev/null 2>&1
npx --no-install tsc --skipLibCheck --target es2020 --lib es2020,dom \
  --outDir "$BUILD_DIR/writer" office_scripts/custsol_append_rawdata_eur.ts >/dev/null 2>&1

if [[ ! -f "$BUILD_DIR/reader/custsol_read_daily.js" || ! -f "$BUILD_DIR/writer/custsol_append_rawdata_eur.js" ]]; then
  echo "compilation produced no output — check tsc errors:" >&2
  npx --no-install tsc --skipLibCheck --target es2020 --lib es2020,dom \
    --outDir "$BUILD_DIR/reader" office_scripts/custsol_read_daily.ts
  exit 1
fi

node tests/custsol_test.js "$BUILD_DIR"
