# Snowflake CI/CD Demo

A demo repository showing how to manage Snowflake objects and Airflow DAGs from a single Git repo with automated CI/CD.

## Structure

```
.
├── snowflake/          # SQL migration files (deployed by schemachange)
├── dags/               # Airflow DAG Python files (synced to Airflow)
├── tests/              # SQL validation scripts
└── .github/workflows/  # GitHub Actions CI/CD pipeline
```

## How It Works

1. **On Pull Request**: GitHub Actions validates SQL file naming and runs a schemachange dry-run
2. **On Merge to Main**: GitHub Actions deploys new SQL migrations to Snowflake via schemachange

## Tools Used

- **[schemachange](https://github.com/Snowflake-Labs/schemachange)** — Snowflake-native database migration tool
- **GitHub Actions** — CI/CD automation
- **Snowflake Git Repository** — Native Git integration for executing scripts from within Snowflake
