-- Migration: V1.1.0
-- Description: Create CUSTOMERS table in RAW schema
-- Author: Peter Pham
-- Deployed via: schemachange + GitHub Actions

CREATE TABLE IF NOT EXISTS RAW.CUSTOMERS (
    customer_id     INTEGER AUTOINCREMENT,
    first_name      VARCHAR(100) NOT NULL,
    last_name       VARCHAR(100) NOT NULL,
    created_at      TIMESTAMP_NTZ DEFAULT CURRENT_TIMESTAMP(),
    PRIMARY KEY (customer_id)
);
