#!/bin/bash
# Run all test suites sequentially and report totals
cd "$(dirname "$0")/.."

echo "════════════════════════════════════════════════════════════"
echo "  Running ALL test suites"
echo "════════════════════════════════════════════════════════════"

TOTAL_EXIT=0
SUITE_COUNT=0
PASS_COUNT=0

run_suite() {
  local label="$1"
  local cmd="$2"
  SUITE_COUNT=$((SUITE_COUNT + 1))
  echo ""
  echo "▸ ${SUITE_COUNT}. ${label}"
  if eval "$cmd"; then
    PASS_COUNT=$((PASS_COUNT + 1))
  else
    TOTAL_EXIT=1
  fi
}

# Unit tests
run_suite "Unit: Overtime"              "npx tsx tests/overtime.test.ts"
run_suite "Unit: Pricing"               "npx tsx tests/pricing.test.ts"
run_suite "Unit: Operational Day"        "npx tsx tests/opDay.test.ts"
run_suite "Unit: Billing Remap"          "npx tsx tests/billingRemap.test.ts"
run_suite "Unit: Marking Aggregation"    "npx tsx tests/markingAggregation.test.ts"
run_suite "Unit: Doc Lifecycle"          "npx tsx tests/docLifecycle.test.ts"

# Integration tests (require running server + DB)
run_suite "Integration: CRUD"            "npx tsx tests/crud_integration.ts"
run_suite "Integration: RBAC"            "npx tsx tests/rbac_tests.ts"
run_suite "Integration: Tenant Isolation" "npx tsx tests/tenant_isolation_tests.ts"
run_suite "Integration: Validation"      "npx tsx tests/validation_tests.ts"
run_suite "Integration: Edge Cases"      "npx tsx tests/edge_case_tests.ts"
run_suite "Integration: Business Logic"  "npx tsx tests/business_logic_tests.ts"

echo ""
echo "════════════════════════════════════════════════════════════"
echo "  ${PASS_COUNT}/${SUITE_COUNT} suites passed"
if [ $TOTAL_EXIT -eq 0 ]; then
  echo "  ALL SUITES PASSED ✓"
else
  echo "  SOME SUITES HAD FAILURES (see above)"
fi
echo "════════════════════════════════════════════════════════════"

exit $TOTAL_EXIT
