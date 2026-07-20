-- Migration: V1.2.0
-- Description: Create a simple ORDERS table and a customer orders view
-- Author: Peter Pham

CREATE TABLE IF NOT EXISTS RAW.ORDERS (
    order_id        INTEGER AUTOINCREMENT,
    customer_id     INTEGER NOT NULL,
    order_total     NUMBER(12,2),
    order_date      DATE DEFAULT CURRENT_DATE(),
    PRIMARY KEY (order_id),
    FOREIGN KEY (customer_id) REFERENCES RAW.CUSTOMERS(customer_id)
);

CREATE OR REPLACE VIEW RAW.CUSTOMER_ORDERS AS
SELECT
    c.customer_id,
    c.first_name || ' ' || c.last_name AS customer_name,
    o.order_id,
    o.order_total,
    o.order_date
FROM RAW.CUSTOMERS c
JOIN RAW.ORDERS o ON c.customer_id = o.customer_id;
