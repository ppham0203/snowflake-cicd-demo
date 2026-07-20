-- Validation: confirm CUSTOMERS table exists and has expected columns
-- This runs during CI to verify the migration succeeded

SELECT
    column_name,
    data_type
FROM CICD_DEMO.INFORMATION_SCHEMA.COLUMNS
WHERE table_schema = 'RAW'
  AND table_name = 'CUSTOMERS'
ORDER BY ordinal_position;
