#!/usr/bin/env bash
set -euo pipefail

MODE="${1:---local}"
BASE_URL="${2:-https://natt4witsfz.github.io/O83_Work-Alert}"
BASE_URL="${BASE_URL%/}"
TMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TMP_DIR"' EXIT

fail() {
  echo "::error::$*" >&2
  exit 1
}

check_file() {
  local file="$1"
  [[ -f "$file" ]] || fail "Missing required file: $file"
}

if [[ "$MODE" == "--local" ]]; then
  check_file "index.html"
  check_file "404.html"
  check_file ".nojekyll"
  check_file "display/index.html"
  cmp -s index.html 404.html || fail "index.html and 404.html are out of sync"
  grep -q '<title>ตารางเข้างาน' index.html || fail "index.html title marker missing"
  grep -q 'function loadStaffJsonp' index.html || fail "GAS loader missing"
  grep -q 'หมดเวลา.*บันทึก' index.html || fail "submit timeout guard missing"
  grep -q 'location.replace' display/index.html || fail "display redirect missing"
  echo "Local Pages validation passed"
  exit 0
fi

[[ "$MODE" == "--live" ]] || fail "Usage: $0 --local | --live [BASE_URL]"

check_live() {
  local path="$1"
  local expected="$2"
  local body="$TMP_DIR/body-$(echo "$path" | tr '/?' '__').html"
  local status
  status="$(curl -L --max-time 30 -sS -o "$body" -w '%{http_code}' "${BASE_URL}${path}")" || fail "Request failed: ${BASE_URL}${path}"
  [[ "$status" == "$expected" ]] || fail "Expected HTTP $expected for ${BASE_URL}${path}, got $status"
  if grep -qi 'Page not found.*GitHub Pages' "$body"; then
    fail "GitHub Pages default 404 body detected at ${BASE_URL}${path}"
  fi
  grep -q '<title>ตารางเข้างาน' "$body" || fail "App title marker missing at ${BASE_URL}${path}"
}

check_live "/" "200"
check_live "/index.html" "200"
check_live "/404.html" "200"
check_live "/display/" "200" "redirect"
echo "Live Pages validation passed for $BASE_URL"
