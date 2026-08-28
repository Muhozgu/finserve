{{ config(
    materialized='view'
) }}

SELECT
    *
FROM {{ ref('risk_assessments') }}