#!/bin/bash
# ==============================================================================
# Comprehensive automated test runner for cw_logs_exporter.sh
# Tests all authentication, authorization, switchback, and excel output paths.
# ==============================================================================

set -uo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TEST_DIR="${BASE_DIR}/tests"
MOCK_BIN="${TEST_DIR}/mock_bin"

mkdir -p "${MOCK_BIN}"

# Create a mock 'aws' command in mock_bin
cat << 'EOF' > "${MOCK_BIN}/aws"
#!/bin/bash
# Mock AWS CLI for testing all scenarios

if [ "$1" = "sts" ] && [ "$2" = "get-caller-identity" ]; then
    if [ "${MOCK_SCENARIO:-}" = "unauthenticated" ]; then
        echo "An error occurred (ExpiredToken) when calling the GetCallerIdentity operation: The security token included in the request is expired" >&2
        exit 254
    else
        echo '{"UserId": "AROAEXAMPLE:session", "Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/TestRole/session"}'
        exit 0
    fi
fi

if [ "$1" = "logs" ] && [ "$2" = "describe-log-groups" ]; then
    if [ "${MOCK_SCENARIO:-}" = "permission_denied" ]; then
        echo "An error occurred (AccessDeniedException) when calling the DescribeLogGroups operation: User is not authorized to perform: logs:DescribeLogGroups on resource: *" >&2
        exit 254
    elif [ "${MOCK_SCENARIO:-}" = "switchback_success" ]; then
        # If switched by switchback script, succeed!
        if [ "${AWS_MOCK_SWITCHED:-0}" = "1" ]; then
            echo '{"logGroups": [{"logGroupName": "/aws/lambda/auth-service"}]}'
            exit 0
        else
            echo "An error occurred (AccessDeniedException) when calling the DescribeLogGroups operation" >&2
            exit 254
        fi
    elif [ "${MOCK_SCENARIO:-}" = "switchback_still_denied" ]; then
        echo "An error occurred (AccessDeniedException) when calling the DescribeLogGroups operation" >&2
        exit 254
    else
        echo '{"logGroups": [{"logGroupName": "/aws/lambda/auth-service"}]}'
        exit 0
    fi
fi

if [ "$1" = "logs" ] && [ "$2" = "filter-log-events" ]; then
    echo '{"events": [{"logStreamName": "stream1", "timestamp": 1791061200000, "message": "TEST LOG", "eventId": "123"}]}'
    exit 0
fi

echo "Mock aws called with: $@"
exit 0
EOF

chmod +x "${MOCK_BIN}/aws"
chmod +x "${TEST_DIR}/mock_switchback.sh"

export PATH="${MOCK_BIN}:${PATH}"

FAILED_TESTS=0
TOTAL_TESTS=0

run_test() {
    local desc="$1"
    local expected_code="$2"
    shift 2
    TOTAL_TESTS=$((TOTAL_TESTS + 1))
    echo "----------------------------------------------------------------"
    echo "[TEST ${TOTAL_TESTS}] ${desc}"
    
    set +e
    OUTPUT=$("$@" 2>&1)
    ACTUAL_CODE=$?
    set -e

    if [ ${ACTUAL_CODE} -eq ${expected_code} ]; then
        echo "  --> RESULT: PASS (Exit code ${ACTUAL_CODE} as expected)"
        echo "  --- Output Captured ---"
        echo "${OUTPUT}" | sed 's/^/      /'
    else
        echo "  --> RESULT: FAIL! Expected code ${expected_code}, got ${ACTUAL_CODE}"
        echo "Output was:"
        echo "${OUTPUT}"
        FAILED_TESTS=$((FAILED_TESTS + 1))
    fi
}

echo "================================================================"
echo "Starting Automated Tests for cw_logs_exporter.sh"
echo "================================================================"

# Test 1: Unauthenticated scenario -> Exit 1, warning message with 'aws login --remote'
export MOCK_SCENARIO="unauthenticated"
run_test "Unauthenticated session detection" 1 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" -b "1h"

# Test 2: Permission denied, auto-switchback disabled -> Exit 2, warning message to switch back
export MOCK_SCENARIO="permission_denied"
run_test "Permission denied with switchback warning" 2 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" -b "1h" -s "${TEST_DIR}/mock_switchback.sh"

# Test 3: Permission denied, auto-switchback enabled -> Auto source switchback script, Exit 0
export MOCK_SCENARIO="switchback_success"
run_test "Auto-switchback execution and recovery" 0 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" -b "1h" -a -s "${TEST_DIR}/mock_switchback.sh" -o "${TEST_DIR}/test_sw_output.xlsx"

# Test 4: Auto-switchback specified but script does not exist -> Exit 3
export MOCK_SCENARIO="switchback_success"
run_test "Switchback script missing detection" 3 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" -b "1h" -a -s "${TEST_DIR}/non_existent_switchback.sh"

# Test 5: Auto-switchback executed but permissions still denied -> Exit 4
export MOCK_SCENARIO="switchback_still_denied"
run_test "Switchback executed but still denied" 4 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" -b "1h" -a -s "${TEST_DIR}/mock_switchback.sh"

# Test 6: Full end-to-end with mock data
export MOCK_SCENARIO=""
run_test "End-to-end execution with mock JSON and Excel generation" 0 \
    bash "${BASE_DIR}/cw_logs_exporter.sh" --mock-data "${TEST_DIR}/mock_cw_data.json" -b "1h" -H "ERROR,WARN" -o "${TEST_DIR}/test_e2e.xlsx"

# Cleanup mock bin
rm -rf "${MOCK_BIN}"

echo "================================================================"
echo "Automated Tests Summary: $((TOTAL_TESTS - FAILED_TESTS))/${TOTAL_TESTS} Passed."
if [ ${FAILED_TESTS} -eq 0 ]; then
    echo "ALL TESTS PASSED SUCCESSFULLY!"
    exit 0
else
    echo "SOME TESTS FAILED!"
    exit 1
fi
