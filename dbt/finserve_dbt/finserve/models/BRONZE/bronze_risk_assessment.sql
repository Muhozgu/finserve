{{ config(
    materialized='view',
    schema='BRONZE'
) }}

SELECT
    *
FROM {{ source('bronze', 'risk_assessments') }}