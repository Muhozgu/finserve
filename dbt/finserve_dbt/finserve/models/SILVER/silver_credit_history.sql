{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized credit history data with trend analysis'
    )
}}

WITH bronze_credit_history AS (
    SELECT * FROM {{ ref('bronze_credit_history') }}
),

-- Clean numeric fields (credit_score)
numeric_cleaned AS (
    SELECT
        *,
        -- Clean credit_score (should be between 300-850)
        CASE 
            WHEN TRY_TO_NUMBER(credit_score) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(credit_score) < 300 OR TRY_TO_NUMBER(credit_score) > 850 THEN NULL
            ELSE TRY_TO_NUMBER(credit_score)
        END AS credit_score_clean,
        -- Flag for out-of-range credit score
        CASE 
            WHEN TRY_TO_NUMBER(credit_score) IS NOT NULL 
                AND (TRY_TO_NUMBER(credit_score) < 300 OR TRY_TO_NUMBER(credit_score) > 850) THEN TRUE
            ELSE FALSE
        END AS credit_score_out_of_range,
        -- Flag for missing credit_score
        CASE 
            WHEN credit_score IS NULL OR TRIM(credit_score) = '' THEN TRUE
            ELSE FALSE
        END AS credit_score_missing,
        
        -- Clean total_outstanding_debt (should be non-negative)
        CASE 
            WHEN TRY_TO_NUMBER(total_outstanding_debt) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(total_outstanding_debt) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(total_outstanding_debt)
        END AS total_outstanding_debt_clean,
        -- Flag for missing/negative debt
        CASE 
            WHEN total_outstanding_debt IS NULL OR TRIM(total_outstanding_debt) = '' 
                OR TRY_TO_NUMBER(total_outstanding_debt) < 0 THEN TRUE
            ELSE FALSE
        END AS total_debt_missing,
        
        -- Clean credit_utilization (should be between 0-1)
        CASE 
            WHEN TRY_TO_NUMBER(credit_utilization) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(credit_utilization) < 0 OR TRY_TO_NUMBER(credit_utilization) > 1 THEN NULL
            ELSE TRY_TO_NUMBER(credit_utilization)
        END AS credit_utilization_clean,
        -- Flag for invalid credit utilization
        CASE 
            WHEN TRY_TO_NUMBER(credit_utilization) IS NOT NULL 
                AND (TRY_TO_NUMBER(credit_utilization) < 0 OR TRY_TO_NUMBER(credit_utilization) > 1) THEN TRUE
            ELSE FALSE
        END AS credit_utilization_invalid
    FROM bronze_credit_history
),

-- Clean integer fields
integer_cleaned AS (
    SELECT
        *,
        -- Clean number_of_open_accounts
        CASE 
            WHEN TRY_TO_NUMBER(number_of_open_accounts) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(number_of_open_accounts) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(number_of_open_accounts)
        END AS number_of_open_accounts_clean,
        
        -- Clean number_of_closed_accounts
        CASE 
            WHEN TRY_TO_NUMBER(number_of_closed_accounts) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(number_of_closed_accounts) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(number_of_closed_accounts)
        END AS number_of_closed_accounts_clean,
        
        -- Calculate total accounts
        CASE 
            WHEN number_of_open_accounts_clean IS NOT NULL AND number_of_closed_accounts_clean IS NOT NULL
            THEN number_of_open_accounts_clean + number_of_closed_accounts_clean
            ELSE NULL
        END AS total_accounts,
        
        -- Clean previous_defaults
        CASE 
            WHEN TRY_TO_NUMBER(previous_defaults) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(previous_defaults) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(previous_defaults)
        END AS previous_defaults_clean,
        
        -- Clean previous_late_payments
        CASE 
            WHEN TRY_TO_NUMBER(previous_late_payments) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(previous_late_payments) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(previous_late_payments)
        END AS previous_late_payments_clean,
        
        -- Flag for high risk (multiple defaults)
        CASE 
            WHEN previous_defaults_clean >= 2 THEN TRUE
            ELSE FALSE
        END AS multiple_defaults,
        
        -- Flag for chronic lateness
        CASE 
            WHEN previous_late_payments_clean >= 3 THEN TRUE
            ELSE FALSE
        END AS chronic_late_payer
    FROM numeric_cleaned
),

-- Date cleaning
date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(record_date) AS record_date_clean,
        -- Flag for future dates
        CASE 
            WHEN TRY_TO_DATE(record_date) > CURRENT_DATE() THEN TRUE
            ELSE FALSE
        END AS record_date_future,
        -- Flag for very old dates
        CASE 
            WHEN TRY_TO_DATE(record_date) < '2010-01-01' THEN TRUE
            ELSE FALSE
        END AS record_date_old
    FROM integer_cleaned
),

-- Create credit score bands
credit_bands AS (
    SELECT
        *,
        CASE 
            WHEN credit_score_clean >= 750 THEN 'Excellent'
            WHEN credit_score_clean >= 700 THEN 'Good'
            WHEN credit_score_clean >= 650 THEN 'Fair'
            WHEN credit_score_clean >= 600 THEN 'Poor'
            WHEN credit_score_clean IS NOT NULL THEN 'Very Poor'
            ELSE 'Unknown'
        END AS credit_score_band,
        -- Utilization bands
        CASE 
            WHEN credit_utilization_clean <= 0.10 THEN 'Very Low (<10%)'
            WHEN credit_utilization_clean <= 0.30 THEN 'Low (10-30%)'
            WHEN credit_utilization_clean <= 0.50 THEN 'Moderate (30-50%)'
            WHEN credit_utilization_clean <= 0.75 THEN 'High (50-75%)'
            WHEN credit_utilization_clean IS NOT NULL THEN 'Very High (>75%)'
            ELSE 'Unknown'
        END AS utilization_band,
        -- Debt level bands
        CASE 
            WHEN total_outstanding_debt_clean < 1000 THEN 'Very Low (<$1K)'
            WHEN total_outstanding_debt_clean < 5000 THEN 'Low ($1K-$5K)'
            WHEN total_outstanding_debt_clean < 15000 THEN 'Moderate ($5K-$15K)'
            WHEN total_outstanding_debt_clean < 30000 THEN 'High ($15K-$30K)'
            WHEN total_outstanding_debt_clean IS NOT NULL THEN 'Very High (>$30K)'
            ELSE 'Unknown'
        END AS debt_level_band
    FROM date_cleaned
),

-- Final selection
final AS (
    SELECT
        -- IDs
        credit_history_id,
        customer_id,
        
        -- Date
        record_date_clean AS record_date,
        record_date_future,
        record_date_old,
        
        -- Credit score
        credit_score_clean AS credit_score,
        credit_score_band,
        credit_score_out_of_range,
        credit_score_missing,
        
        -- Debt metrics
        total_outstanding_debt_clean AS total_outstanding_debt,
        total_debt_missing,
        debt_level_band,
        credit_utilization_clean AS credit_utilization,
        credit_utilization_invalid,
        utilization_band,
        
        -- Account metrics
        number_of_open_accounts_clean AS number_of_open_accounts,
        number_of_closed_accounts_clean AS number_of_closed_accounts,
        total_accounts,
        
        -- Delinquency metrics
        previous_defaults_clean AS previous_defaults,
        multiple_defaults,
        previous_late_payments_clean AS previous_late_payments,
        chronic_late_payer,
        
        -- Data quality flags
        CASE 
            WHEN credit_score_clean IS NULL 
                OR total_outstanding_debt_clean IS NULL
                OR credit_utilization_clean IS NULL
                OR credit_score_out_of_range = TRUE
                OR credit_utilization_invalid = TRUE
                OR record_date_future = TRUE
            THEN TRUE
            ELSE FALSE
        END AS has_data_quality_issue,
        
        -- Created/updated timestamp
        
    FROM credit_bands
)

SELECT * FROM final