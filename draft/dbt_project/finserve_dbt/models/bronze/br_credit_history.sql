{{ config(
    materialized='view'
) }}

SELECT
    *
FROM {{ ref('credit_history') }}