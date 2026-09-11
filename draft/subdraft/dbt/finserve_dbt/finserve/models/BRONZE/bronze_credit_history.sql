{{ config(
    materialized='view',
    schema='BRONZE'
) }}

SELECT
    *
FROM {{ source('bronze', 'credit_history') }}