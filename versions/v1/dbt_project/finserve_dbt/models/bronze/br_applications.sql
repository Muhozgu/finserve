{{ config(
    materialized='view'
) }}

SELECT
    *
FROM {{ ref('applications') }}
