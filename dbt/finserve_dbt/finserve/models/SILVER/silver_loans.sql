{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized loans data with loan performance analysis'
    )
}}

WITH bronze_loans AS (
    SELECT * FROM {{ ref('bronze_loans') }}
),

-- Clean and standardize loan status
status_cleaned AS (
    SELECT
        *,
        -- Standardize loan status
        CASE 
            WHEN UPPER(TRIM(loan_status)) IN ('ACTIVE', 'ACTIVE ') THEN 'Active'
            WHEN UPPER(TRIM(loan_status)) IN ('PAID OFF', 'PAID OFF ', 'PAID', 'PAID ') THEN 'Paid Off'
            WHEN UPPER(TRIM(loan_status)) IN ('DEFAULTED', 'DEFAULTED ', 'DEFAULT', 'DEFAULT ') THEN 'Defaulted'
            WHEN UPPER(TRIM(loan_status)) IN ('DELINQUENT', 'DELINQUENT ', 'DELINQ') THEN 'Delinquent'
            WHEN loan_status IS NULL OR TRIM(loan_status) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(loan_status))
        END AS loan_status_standardized,
        -- Flag for valid status
        CASE 
            WHEN UPPER(TRIM(loan_status)) IN ('ACTIVE', 'ACTIVE ', 'PAID OFF', 'PAID OFF ', 'PAID', 'PAID ',
                                              'DEFAULTED', 'DEFAULTED ', 'DEFAULT', 'DEFAULT ',
                                              'DELINQUENT', 'DELINQUENT ', 'DELINQ') THEN TRUE
            ELSE FALSE
        END AS has_valid_status
    FROM bronze_loans
),

-- Clean numeric fields
numeric_cleaned AS (
    SELECT
        *,
        -- Clean loan_amount (should be positive, handle negative values)
        CASE 
            WHEN TRY_TO_NUMBER(loan_amount) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(loan_amount) < 0 THEN NULL  -- Negative loan amount is invalid
            ELSE TRY_TO_NUMBER(loan_amount)
        END AS loan_amount_clean,
        -- Flag for missing/negative loan_amount
        CASE 
            WHEN loan_amount IS NULL OR TRIM(loan_amount) = '' OR TRY_TO_NUMBER(loan_amount) < 0 THEN TRUE
            ELSE FALSE
        END AS loan_amount_missing,
        
        -- Clean interest_rate (should be between 0-100, handle negative values)
        CASE 
            WHEN TRY_TO_NUMBER(interest_rate) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(interest_rate) < 0 THEN NULL
            WHEN TRY_TO_NUMBER(interest_rate) > 100 THEN NULL  -- Unrealistic interest rate
            ELSE TRY_TO_NUMBER(interest_rate)
        END AS interest_rate_clean,
        -- Flag for missing/invalid interest_rate
        CASE 
            WHEN interest_rate IS NULL OR TRIM(interest_rate) = '' 
                OR TRY_TO_NUMBER(interest_rate) < 0 
                OR TRY_TO_NUMBER(interest_rate) > 100 THEN TRUE
            ELSE FALSE
        END AS interest_rate_invalid,
        
        -- Clean term_months (should be positive)
        CASE 
            WHEN TRY_TO_NUMBER(term_months) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(term_months) <= 0 THEN NULL
            ELSE TRY_TO_NUMBER(term_months)
        END AS term_months_clean,
        -- Flag for missing/negative term_months
        CASE 
            WHEN term_months IS NULL OR TRIM(term_months) = '' OR TRY_TO_NUMBER(term_months) <= 0 THEN TRUE
            ELSE FALSE
        END AS term_months_missing,
        
        -- Clean monthly_payment (should be positive)
        CASE 
            WHEN TRY_TO_NUMBER(monthly_payment) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(monthly_payment) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(monthly_payment)
        END AS monthly_payment_clean,
        -- Flag for missing/negative monthly_payment
        CASE 
            WHEN monthly_payment IS NULL OR TRIM(monthly_payment) = '' OR TRY_TO_NUMBER(monthly_payment) < 0 THEN TRUE
            ELSE FALSE
        END AS monthly_payment_missing
    FROM status_cleaned
),

-- Clean and validate dates
date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(start_date) AS start_date_clean,
        TRY_TO_DATE(maturity_date) AS maturity_date_clean,
        -- Flag for missing start_date
        CASE 
            WHEN start_date IS NULL OR TRIM(start_date) = '' THEN TRUE
            ELSE FALSE
        END AS start_date_missing,
        -- Flag for missing maturity_date
        CASE 
            WHEN maturity_date IS NULL OR TRIM(maturity_date) = '' THEN TRUE
            ELSE FALSE
        END AS maturity_date_missing,
        -- Check if maturity_date is after start_date
        CASE 
            WHEN TRY_TO_DATE(start_date) IS NOT NULL AND TRY_TO_DATE(maturity_date) IS NOT NULL
                AND TRY_TO_DATE(maturity_date) <= TRY_TO_DATE(start_date) THEN TRUE
            ELSE FALSE
        END AS maturity_date_anomaly,
        -- Calculate loan duration in days
        DATEDIFF('day', 
            TRY_TO_DATE(start_date), 
            TRY_TO_DATE(maturity_date)
        ) AS loan_duration_days,
        -- Calculate months remaining (for active loans)
        CASE 
            WHEN loan_status_standardized IN ('Active', 'Delinquent') 
                AND TRY_TO_DATE(maturity_date) IS NOT NULL
            THEN DATEDIFF('month', CURRENT_DATE(), TRY_TO_DATE(maturity_date))
            ELSE NULL
        END AS months_remaining,
        -- Check if loan is past maturity date
        CASE 
            WHEN TRY_TO_DATE(maturity_date) IS NOT NULL 
                AND TRY_TO_DATE(maturity_date) < CURRENT_DATE()
                AND loan_status_standardized NOT IN ('Paid Off', 'Defaulted') THEN TRUE
            ELSE FALSE
        END AS past_maturity
    FROM numeric_cleaned
),

-- Calculate loan performance metrics
performance_metrics AS (
    SELECT
        *,
        -- Calculate estimated total cost of loan (principal + interest)
        CASE 
            WHEN monthly_payment_clean IS NOT NULL AND term_months_clean IS NOT NULL
            THEN monthly_payment_clean * term_months_clean
            ELSE NULL
        END AS estimated_total_cost,
        -- Calculate total interest estimated
        CASE 
            WHEN monthly_payment_clean IS NOT NULL AND term_months_clean IS NOT NULL 
                AND loan_amount_clean IS NOT NULL
            THEN (monthly_payment_clean * term_months_clean) - loan_amount_clean
            ELSE NULL
        END AS estimated_total_interest,
        -- Calculate interest rate band
        CASE 
            WHEN interest_rate_clean < 5 THEN 'Very Low (<5%)'
            WHEN interest_rate_clean < 8 THEN 'Low (5-8%)'
            WHEN interest_rate_clean < 12 THEN 'Medium (8-12%)'
            WHEN interest_rate_clean < 16 THEN 'High (12-16%)'
            WHEN interest_rate_clean IS NOT NULL THEN 'Very High (>16%)'
            ELSE 'Unknown'
        END AS interest_rate_band,
        -- Calculate loan amount band
        CASE 
            WHEN loan_amount_clean < 5000 THEN 'Small (<$5K)'
            WHEN loan_amount_clean < 15000 THEN 'Medium ($5K-$15K)'
            WHEN loan_amount_clean < 30000 THEN 'Large ($15K-$30K)'
            WHEN loan_amount_clean < 50000 THEN 'Very Large ($30K-$50K)'
            WHEN loan_amount_clean IS NOT NULL THEN 'Jumbo (>$50K)'
            ELSE 'Unknown'
        END AS loan_size_band
    FROM date_cleaned
),

-- Final selection
final AS (
    SELECT
        -- IDs
        loan_id,
        application_id,
        customer_id,
        
        -- Loan details
        loan_amount_clean AS loan_amount,
        loan_amount_missing,
        interest_rate_clean AS interest_rate,
        interest_rate_invalid,
        interest_rate_band,
        term_months_clean AS term_months,
        term_months_missing,
        monthly_payment_clean AS monthly_payment,
        monthly_payment_missing,
        
        -- Dates
        start_date_clean AS start_date,
        start_date_missing,
        maturity_date_clean AS maturity_date,
        maturity_date_missing,
        maturity_date_anomaly,
        loan_duration_days,
        months_remaining,
        past_maturity,
        
        -- Status
        loan_status_standardized AS loan_status,
        has_valid_status,
        
        -- Performance metrics
        estimated_total_cost,
        estimated_total_interest,
        loan_size_band,
        
        -- Created/updated timestamps
        created_at,
        updated_at
        
    FROM performance_metrics
)

SELECT * FROM final