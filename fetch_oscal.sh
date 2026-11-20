#!/bin/bash
# Download the NIST SP 800-53 Rev 5 OSCAL catalog and SP 800-53B baseline profiles.
# Usage: ./fetch_oscal.sh [output_dir]   (default: ./oscal)
set -euo pipefail
OUT="${1:-oscal}"
mkdir -p "$OUT"
BASE=https://raw.githubusercontent.com/usnistgov/oscal-content/main/nist.gov/SP800-53/rev5/json
for f in NIST_SP-800-53_rev5_catalog.json \
         NIST_SP-800-53_rev5_LOW-baseline_profile.json \
         NIST_SP-800-53_rev5_MODERATE-baseline_profile.json \
         NIST_SP-800-53_rev5_HIGH-baseline_profile.json \
         NIST_SP-800-53_rev5_PRIVACY-baseline_profile.json; do
  curl -sfL -o "$OUT/$f" "$BASE/$f"
  echo "downloaded $f"
done
