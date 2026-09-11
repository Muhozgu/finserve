{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized applications data'
    )
}}

WITH bronze_applications AS (
    SELECT * FROM {{ ref('bronze_applications') }}
),

status_cleaned AS (
    SELECT
        *,
        UPPER(TRIM(application_status)) AS application_status_clean,
        CASE 
            WHEN UPPER(TRIM(application_status)) IN ('APPROVED', 'APPR', 'APPROVE', 'APPR.', 'APPROVED ') THEN 'Approved'
            WHEN UPPER(TRIM(application_status)) IN ('REJECTED', 'REJECT', 'REJECTED ', 'DECLINED', 'DECL', 'DECLIN', 'DECLINED ') THEN 'Rejected'
            WHEN UPPER(TRIM(application_status)) IN ('PENDING', 'PEND', 'PENDING ') THEN 'Pending'
            WHEN UPPER(TRIM(application_status)) IN ('WITHDRAWN', 'WITHDRAW', 'WITHDRAWN ') THEN 'Withdrawn'
            WHEN UPPER(TRIM(application_status)) IN ('IN REVIEW', 'IN REVIEW ') THEN 'In Review'
            WHEN UPPER(TRIM(application_status)) IN ('CANCELLED', 'CANCEL', 'CANCELLED ') THEN 'Cancelled'
            ELSE UPPER(TRIM(application_status))
        END AS application_status_standardized,
        CASE 
            WHEN UPPER(TRIM(application_status)) IN ('APPROVED', 'APPR', 'APPROVE', 'APPR.',
                                                      'REJECTED', 'REJECT', 'REJECTED ', 'DECLINED', 'DECL', 'DECLIN',
                                                      'PENDING', 'PEND', 'PENDING ',
                                                      'WITHDRAWN', 'WITHDRAW', 'WITHDRAWN ',
                                                      'IN REVIEW', 'IN REVIEW ',
                                                      'CANCELLED', 'CANCEL', 'CANCELLED ') THEN TRUE
            ELSE FALSE
        END AS has_valid_status
    FROM bronze_applications
),

numeric_cleaned AS (
    SELECT
        *,
        TRY_TO_NUMBER(REPLACE(requested_amount, ',', '')) AS requested_amount_clean,
        CASE 
            WHEN requested_amount IS NULL OR TRIM(requested_amount) = '' THEN TRUE 
            ELSE FALSE 
        END AS requested_amount_missing
    FROM status_cleaned
),

date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(application_date) AS application_date_clean,
        TRY_TO_DATE(decision_date) AS decision_date_clean,
        CASE 
            WHEN TRY_TO_DATE(application_date) IS NOT NULL 
                AND TRY_TO_DATE(decision_date) IS NOT NULL
                AND TRY_TO_DATE(decision_date) < TRY_TO_DATE(application_date) THEN TRUE
            ELSE FALSE
        END AS decision_date_anomaly,
        DATEDIFF('day', 
            TRY_TO_DATE(application_date), 
            TRY_TO_DATE(decision_date)
        ) AS processing_days
    FROM numeric_cleaned
),

channel_cleaned AS (
    SELECT
        *,
        CASE 
            WHEN application_channel IS NULL OR TRIM(application_channel) = '' THEN 'Unknown'
            ELSE UPPER(TRIM(application_channel))
        END AS application_channel_standardized
    FROM date_cleaned
),

final AS (
    SELECT
        application_id,
        customer_id,
        application_date_clean AS application_date,
        decision_date_clean AS decision_date,
        processing_days,
        decision_date_anomaly,
        requested_amount_clean AS requested_amount,
        requested_amount_missing,
        CASE 
            WHEN loan_purpose IS NULL OR TRIM(loan_purpose) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(loan_purpose))
        END AS loan_purpose,
        application_status_standardized AS application_status,
        CASE 
            WHEN application_status_standardized IN ('Approved', 'Rejected', 'Withdrawn', 'Cancelled') THEN TRUE
            ELSE FALSE
        END AS is_final_decision,
        application_channel_standardized AS application_channel,
        has_valid_status
    FROM channel_cleaned
)

SELECT * FROM final