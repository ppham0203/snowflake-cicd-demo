# Plan: End-to-End Git + CI/CD Demo Setup

## Overview

Build a working demo that deploys Snowflake schema changes and an Airflow DAG from a GitHub repo using GitHub Actions. The demo targets account `SFSENORTHAMERICA-PPHAM_AWS1` with warehouse `airflow_wh`.

---

## Phase 1: Create the Demo Files Locally

All files go in `/Users/peterpham/projects/Enablement/Git_CICD/`.

```
.
├── snowflake/
│   ├── V1.0.0__create_raw_schema.sql
│   ├── V1.1.0__create_customers_table.sql
│   └── V1.2.0__create_orders_view.sql
├── dags/
│   └── ingest_customers_dag.py   (sample Airflow DAG — doesn't run, just shows the pattern)
├── tests/
│   └── validate_customers.sql    (simple SELECT to verify the table exists)
├── .github/
│   └── workflows/
│       └── deploy.yml            (GitHub Actions workflow: lint SQL + deploy via schemachange)
├── schemachange-config.yml       (schemachange config file)
└── README.md                     (optional, brief explanation)
```

**Key decisions:**
- **schemachange** will be the migration tool (open-source Python tool by Snowflake)
- **GitHub Actions** will install schemachange via `pip` and run it on merge to `main`
- **DAG deployment** will be simulated — the workflow echoes "DAG sync would go here" (no live Airflow)
- **Authentication**: We'll use Snowflake username + password stored in GitHub Secrets (simplest for a demo)

---

## Phase 2: Initialize Git and Push to GitHub

1. `git init` the project
2. `gh repo create peter-pham_snow/snowflake-cicd-demo --public --source=. --push`
3. Confirm the repo is visible at github.com

---

## Phase 3: Create Snowflake Objects for the Demo

Run these in Snowflake:
```sql
CREATE DATABASE IF NOT EXISTS CICD_DEMO;
CREATE SCHEMA IF NOT EXISTS CICD_DEMO.RAW;
CREATE WAREHOUSE IF NOT EXISTS AIRFLOW_WH
  WITH WAREHOUSE_SIZE = 'XSMALL' AUTO_SUSPEND = 60 AUTO_RESUME = TRUE;
```

schemachange will auto-create its tracking table (`CHANGE_HISTORY`) on first run.

---

## Phase 4: Snowflake Git Repository Integration

This lets Snowflake *pull* files directly from the GitHub repo (useful for executing scripts from within Snowflake, independent of CI/CD).

**Steps:**
1. Generate a GitHub Personal Access Token (PAT) with `repo` scope
2. Store it in Snowflake:
   ```sql
   CREATE OR REPLACE SECRET CICD_DEMO.RAW.GITHUB_PAT
     TYPE = PASSWORD
     USERNAME = 'peter-pham_snow'
     PASSWORD = '<your-pat>';
   ```
3. Create an API integration:
   ```sql
   CREATE OR REPLACE API INTEGRATION GITHUB_INTEGRATION
     API_PROVIDER = GIT_HTTPS_API
     API_ALLOWED_PREFIXES = ('https://github.com/peter-pham_snow/')
     ALLOWED_AUTHENTICATION_SECRETS = (CICD_DEMO.RAW.GITHUB_PAT)
     ENABLED = TRUE;
   ```
4. Create the Git Repository:
   ```sql
   CREATE OR REPLACE GIT REPOSITORY CICD_DEMO.RAW.CICD_REPO
     API_INTEGRATION = GITHUB_INTEGRATION
     GIT_CREDENTIALS = CICD_DEMO.RAW.GITHUB_PAT
     ORIGIN = 'https://github.com/peter-pham_snow/snowflake-cicd-demo.git';
   ```
5. Test: `ALTER GIT REPOSITORY CICD_DEMO.RAW.CICD_REPO FETCH;`
6. List files: `SHOW GIT BRANCHES IN CICD_DEMO.RAW.CICD_REPO;`

---

## Phase 5: Add GitHub Actions Secrets

On the GitHub repo, store these secrets (via `gh secret set`):

| Secret Name | Value |
|---|---|
| `SNOWFLAKE_ACCOUNT` | `SFSENORTHAMERICA-PPHAM_AWS1` |
| `SNOWFLAKE_USER` | `PEPHAM` |
| `SNOWFLAKE_PASSWORD` | Your Snowflake password (or we can use key-pair auth) |
| `SNOWFLAKE_WAREHOUSE` | `AIRFLOW_WH` |
| `SNOWFLAKE_DATABASE` | `CICD_DEMO` |
| `SNOWFLAKE_ROLE` | `ACCOUNTADMIN` (or a dedicated deploy role) |

---

## Phase 6: Test the Full Loop

1. Create `snowflake/V2.0.0__add_email_column.sql` on a feature branch
2. Push and open a PR
3. CI runs: validates SQL syntax
4. Merge PR to main
5. GitHub Actions runs schemachange → deploys the migration to Snowflake
6. Verify in Snowflake: `DESCRIBE TABLE CICD_DEMO.RAW.CUSTOMERS;` shows the new column

---

## What You'll Learn Along the Way

- How to create a GitHub repo from the terminal
- How GitHub Actions YAML files work (triggers, jobs, steps)
- How to store secrets safely (not in code)
- How schemachange tracks migrations
- How Snowflake's native Git Repository feature works
- How to simulate DAG deployment in the same pipeline

---

## Security Notes

- The demo uses password auth for simplicity. In production, use key-pair auth or OAuth.
- GitHub Secrets are encrypted and not visible in logs — safe for demos.
- For a real customer engagement, you'd use a dedicated service account role (not ACCOUNTADMIN).
