-- Migration: V2.0.0
-- Description: Add email column to CUSTOMERS table
-- Author: Peter Pham
-- NOTE: Uses IF NOT EXISTS so this is safe to re-run if needed
ALTER TABLE RAW.CUSTOMERS ADD COLUMN IF NOT EXISTS email VARCHAR(255);
