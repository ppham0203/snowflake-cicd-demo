#!/bin/bash
# =============================================================================
# demo-reset.sh — Reset the CI/CD demo to a clean state
#
# Run this BEFORE each demo practice or live run.
# It removes the V2.0.0 migration from both Snowflake and Git so the
# full demo flow (create → PR → merge → deploy → verify) works from scratch.
#
# Usage:
#   chmod +x scripts/demo-reset.sh   (first time only)
#   ./scripts/demo-reset.sh
# =============================================================================
set -e

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO_ROOT"

DEMO_FILE="snowflake/V2.0.0__add_email_to_customers.sql"
DEMO_BRANCH="feature/add-email-column"
DEMO_VERSION="2.0.0"

echo ""
echo "══════════════════════════════════════════════"
echo "  CI/CD Demo Reset"
echo "══════════════════════════════════════════════"
echo ""

# ── Step 1: Ensure we're on main and up to date ────────────────────────────
echo "1/5  Switching to main and pulling latest..."
git checkout main
git pull origin main --quiet
echo "     ✓ On main, up to date"

# ── Step 2: Delete the demo migration file if it exists ───────────────────
if [ -f "$DEMO_FILE" ]; then
  echo "2/5  Removing demo migration file..."
  git rm -f "$DEMO_FILE" 2>/dev/null || rm -f "$DEMO_FILE"
  git add -A
  git commit -m "Reset: remove V2.0.0 demo migration for re-run" --quiet
  git push origin main --quiet
  echo "     ✓ Migration file removed and pushed"
else
  echo "2/5  Migration file not present — nothing to remove ✓"
fi

# ── Step 3: Delete the local feature branch if it exists ──────────────────
echo "3/5  Cleaning up feature branch..."
if git show-ref --verify --quiet refs/heads/$DEMO_BRANCH; then
  git branch -D "$DEMO_BRANCH" 2>/dev/null
  echo "     ✓ Local branch '$DEMO_BRANCH' deleted"
else
  echo "     Branch '$DEMO_BRANCH' not found locally — skipping ✓"
fi

# Delete remote feature branch if it exists
if git ls-remote --heads origin "$DEMO_BRANCH" | grep -q "$DEMO_BRANCH"; then
  git push origin --delete "$DEMO_BRANCH" --quiet 2>/dev/null || true
  echo "     ✓ Remote branch '$DEMO_BRANCH' deleted"
fi

# ── Step 4: Close any open PRs for the demo branch ────────────────────────
echo "4/5  Closing any open PRs for the demo branch..."
OPEN_PRS=$(gh pr list --repo ppham0203/snowflake-cicd-demo --head "$DEMO_BRANCH" --json number -q '.[].number' 2>/dev/null || echo "")
if [ -n "$OPEN_PRS" ]; then
  for PR in $OPEN_PRS; do
    gh pr close "$PR" --repo ppham0203/snowflake-cicd-demo --comment "Reset for demo re-run" --quiet
    echo "     ✓ Closed PR #$PR"
  done
else
  echo "     No open PRs — nothing to close ✓"
fi

# ── Step 5: Print the Snowflake reset SQL (must run manually in Snowsight) ─
echo "5/5  Snowflake state needs manual reset."
echo ""
echo "     ┌─────────────────────────────────────────────────────────────┐"
echo "     │  Run this SQL in Snowsight to complete the reset:           │"
echo "     │                                                             │"
echo "     │  ALTER TABLE CICD_DEMO.RAW.CUSTOMERS                        │"
echo "     │    DROP COLUMN IF EXISTS email;                             │"
echo "     │                                                             │"
echo "     │  DELETE FROM CICD_DEMO.RAW.CHANGE_HISTORY                   │"
echo "     │    WHERE VERSION = '$DEMO_VERSION';                           │"
echo "     └─────────────────────────────────────────────────────────────┘"
echo ""

# ── Done ────────────────────────────────────────────────────────────────────
echo "══════════════════════════════════════════════"
echo "  Reset complete."
echo ""
echo "  Git state:      ✓ Clean (on main, no demo branch)"
echo "  Snowflake:      ⚠  Run the SQL above in Snowsight"
echo ""
echo "  You can now run the full demo from scratch:"
echo "  → Create V2.0.0__add_email_to_customers.sql"
echo "  → git checkout -b feature/add-email-column"
echo "  → git add, commit, push"
echo "  → gh pr create, merge"
echo "  → Watch GitHub Actions deploy to Snowflake"
echo "══════════════════════════════════════════════"
echo ""
