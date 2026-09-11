{{ config(materialized='table') }}

WITH cleaned AS(
    SELECT
        CUSTOMER_ID,
        FIRST_NAME,
        LAST_NAME,
        DATE_OF_BIRTH,
        GENDER,
        COUNTRY,
        CITY,
        EMPLOYMENT_STATUS,
        EMPLOYMENT_LENGTH_YEARS,
        ANNUAL_INCOME,
        MONTHLY_INCOME,

),

deduplicated as (
    SELECT *,
        ROW_NUMBER() OVER (PARTITION BY transaction_id ORDER BY order_date DESC) AS row_num
    FROM cleaned
)


SELECT * FROM deduplicated
WHERE row_num = 1