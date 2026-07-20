"""
Sample Airflow DAG: ingest_customers

This DAG demonstrates how Airflow calls Snowflake via the Snowflake operator.
It lives in the same repo as the Snowflake migrations — one repo, one pipeline.

In a real setup, CI/CD syncs this file to the Airflow DAG folder
(S3 for MWAA, GCS for Cloud Composer, or a mounted volume for local).
"""
from datetime import datetime, timedelta

from airflow import DAG
from airflow.providers.snowflake.operators.snowflake import SnowflakeOperator

default_args = {
    "owner": "data-engineering",
    "depends_on_past": False,
    "retries": 1,
    "retry_delay": timedelta(minutes=5),
}

with DAG(
    dag_id="ingest_customers",
    description="Load new customers from source into Snowflake RAW schema",
    default_args=default_args,
    start_date=datetime(2026, 1, 1),
    schedule_interval="@daily",
    catchup=False,
    tags=["snowflake", "ingestion", "demo"],
) as dag:

    load_customers = SnowflakeOperator(
        task_id="load_customers",
        snowflake_conn_id="snowflake_default",
        sql="""
            INSERT INTO CICD_DEMO.RAW.CUSTOMERS (first_name, last_name)
            SELECT first_name, last_name
            FROM CICD_DEMO.RAW.CUSTOMERS_STAGING
            WHERE NOT EXISTS (
                SELECT 1 FROM CICD_DEMO.RAW.CUSTOMERS c
                WHERE c.first_name = CUSTOMERS_STAGING.first_name
                  AND c.last_name = CUSTOMERS_STAGING.last_name
            );
        """,
    )

    verify_load = SnowflakeOperator(
        task_id="verify_load",
        snowflake_conn_id="snowflake_default",
        sql="SELECT COUNT(*) AS customer_count FROM CICD_DEMO.RAW.CUSTOMERS;",
    )

    load_customers >> verify_load
