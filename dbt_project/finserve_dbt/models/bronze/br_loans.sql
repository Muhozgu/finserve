{{ config(
    materialized='view'
) }}

SELECT
    *
FROM {{ ref('loans') }}