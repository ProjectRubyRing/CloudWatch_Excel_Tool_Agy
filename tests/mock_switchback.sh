#!/bin/bash
# Mock switchback shell provided by another team
# Changes AWS environment to full operation permissions
export AWS_MOCK_SWITCHED=1
export AWS_PROFILE="admin-role-profile"
echo "[MOCK_SWITCHBACK] Switched AWS role/profile to 'admin-role-profile'."
