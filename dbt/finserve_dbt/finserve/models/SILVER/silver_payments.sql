{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized payments data with payment behavior analysis'
    )
}}

WITH bronze_payments AS (
    SELECT * FROM {{ ref('bronze_payments') }}
),

-- Clean and standardize payment status
status_cleaned AS (
    SELECT
        *,
        -- Standardize payment status
        CASE 
            WHEN UPPER(TRIM(payment_status)) IN ('ON-TIME', 'ON TIME', 'ONTIME', 'ON-TIME ', 'OK', 'OK ', 'ON TIME ') THEN 'On-Time'
            WHEN UPPER(TRIM(payment_status)) IN ('LATE', 'LATE ') THEN 'Late'
            WHEN UPPER(TRIM(payment_status)) IN ('MISSED', 'MISSED ', 'NSF', 'FAILED') THEN 'Missed'
            WHEN UPPER(TRIM(payment_status)) IN ('PARTIAL', 'PART PAYMENT', 'PART PAYMENT ', 'PARTIAL ') THEN 'Partial'
            WHEN UPPER(TRIM(payment_status)) IN ('DELAYED', 'DELAYED ') THEN 'Delayed'
            WHEN UPPER(TRIM(payment_status)) IN ('ON TIME', 'ON TIME ') THEN 'On-Time'  -- Duplicate for safety
            WHEN payment_status IS NULL OR TRIM(payment_status) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(payment_status))
        END AS payment_status_standardized,
        -- Flag for valid status
        CASE 
            WHEN UPPER(TRIM(payment_status)) IN ('ON-TIME', 'ON TIME', 'ONTIME', 'ON-TIME ', 'OK', 'OK ', 'ON TIME ',
                                                  'LATE', 'LATE ',
                                                  'MISSED', 'MISSED ', 'NSF', 'FAILED',
                                                  'PARTIAL', 'PART PAYMENT', 'PART PAYMENT ', 'PARTIAL ',
                                                  'DELAYED', 'DELAYED ',
                                                  'ON TIME', 'ON TIME ') THEN TRUE
            ELSE FALSE
        END AS has_valid_status
    FROM bronze_payments
),

-- Clean numeric fields
numeric_cleaned AS (
    SELECT
        *,
        -- Clean amount_due (should be positive)
        CASE 
            WHEN TRY_TO_NUMBER(amount_due) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(amount_due) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(amount_due)
        END AS amount_due_clean,
        -- Flag for missing/negative amount_due
        CASE 
            WHEN amount_due IS NULL OR TRIM(amount_due) = '' OR TRY_TO_NUMBER(amount_due) < 0 THEN TRUE
            ELSE FALSE
        END AS amount_due_missing,
        
        -- Clean amount_paid (can be 0 for missed payments)
        CASE 
            WHEN TRY_TO_NUMBER(amount_paid) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(amount_paid) < 0 THEN NULL  -- Negative payments are refunds/errors
            ELSE TRY_TO_NUMBER(amount_paid)
        END AS amount_paid_clean,
        -- Flag for missing/negative amount_paid
        CASE 
            WHEN amount_paid IS NULL OR TRIM(amount_paid) = '' OR TRY_TO_NUMBER(amount_paid) < 0 THEN TRUE
            ELSE FALSE
        END AS amount_paid_missing,
        
        -- Calculate payment ratio (what % of amount_due was paid)
        CASE 
            WHEN amount_due_clean IS NOT NULL AND amount_due_clean > 0 
                AND amount_paid_clean IS NOT NULL AND amount_paid_clean >= 0
            THEN ROUND(amount_paid_clean / amount_due_clean, 4)
            ELSE NULL
        END AS payment_ratio,
        
        -- Flag for overpayment
        CASE 
            WHEN amount_paid_clean IS NOT NULL AND amount_due_clean IS NOT NULL
                AND amount_paid_clean > amount_due_clean THEN TRUE
            ELSE FALSE
        END AS is_overpayment
    FROM status_cleaned
),

-- Clean and validate dates
date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(due_date) AS due_date_clean,
        TRY_TO_DATE(payment_date) AS payment_date_clean,
        -- Validate days_late
        CASE 
            WHEN TRY_TO_NUMBER(days_late) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(days_late) < 0 THEN NULL  -- Negative days_late is invalid
            ELSE TRY_TO_NUMBER(days_late)
        END AS days_late_clean,
        -- Flag for missing payment date
        CASE 
            WHEN payment_date IS NULL OR TRIM(payment_date) = '' THEN TRUE
            ELSE FALSE
        END AS payment_date_missing,
        -- Flag for future payment dates
        CASE 
            WHEN TRY_TO_DATE(payment_date) > CURRENT_DATE() THEN TRUE
            ELSE FALSE
        END AS payment_date_future,
        -- Check if payment was made on time (based on days_late)
        CASE 
            WHEN days_late_clean IS NOT NULL AND days_late_clean = 0 THEN TRUE
            WHEN days_late_clean IS NOT NULL AND days_late_clean > 0 THEN FALSE
            ELSE NULL
        END AS was_on_time_by_days,
        -- Check if payment was late (days_late > 0)
        CASE 
            WHEN days_late_clean IS NOT NULL AND days_late_clean > 0 THEN TRUE
            ELSE FALSE
        END AS was_late,
        -- Categorize lateness
        CASE 
            WHEN days_late_clean IS NULL THEN 'Unknown'
            WHEN days_late_clean = 0 THEN 'On-Time'
            WHEN days_late_clean <= 7 THEN '1-7 Days Late'
            WHEN days_late_clean <= 15 THEN '8-15 Days Late'
            WHEN days_late_clean <= 30 THEN '16-30 Days Late'
            WHEN days_late_clean <= 45 THEN '31-45 Days Late'
            ELSE 'Over 45 Days Late'
        END AS late_category,
        -- Flag for very late payments (over 30 days)
        CASE 
            WHEN days_late_clean IS NOT NULL AND days_late_clean > 30 THEN TRUE
            ELSE FALSE
        END AS severely_late
    FROM numeric_cleaned
),

-- Additional payment behavior flags
payment_behavior AS (
    SELECT
        *,
        -- Payment completion flag
        CASE 
            WHEN amount_due_clean IS NOT NULL AND amount_paid_clean IS NOT NULL
                AND amount_paid_clean >= amount_due_clean THEN 'Full'
            WHEN amount_due_clean IS NOT NULL AND amount_paid_clean IS NOT NULL
                AND amount_paid_clean > 0 AND amount_paid_clean < amount_due_clean THEN 'Partial'
            WHEN amount_paid_clean IS NULL OR amount_paid_clean = 0 THEN 'None'
            ELSE 'Unknown'
        END AS payment_completion,
        -- Payment timing behavior
        CASE 
            WHEN payment_date_missing = TRUE AND was_late = FALSE THEN 'Unknown'
            WHEN days_late_clean = 0 THEN 'Early/On-Time'
            WHEN days_late_clean <= 7 THEN 'Slightly Late'
            WHEN days_late_clean <= 30 THEN 'Moderately Late'
            WHEN days_late_clean > 30 THEN 'Significantly Late'
            ELSE 'Unknown'
        END AS payment_timing
    FROM date_cleaned
),

-- Final selection
final AS (
    SELECT
        -- IDs
        payment_id,
        loan_id,
        customer_id,
        
        -- Dates
        due_date_clean AS due_date,
        payment_date_clean AS payment_date,
        payment_date_missing,
        payment_date_future,
        
        -- Amounts
        amount_due_clean AS amount_due,
        amount_paid_clean AS amount_paid,
        amount_due_missing,
        amount_paid_missing,
        payment_ratio,
        is_overpayment,
        
        -- Payment status
        payment_status_standardized AS payment_status,
        has_valid_status,
        
        -- Lateness indicators
        days_late_clean AS days_late,
        was_on_time_by_days,
        was_late,
        late_category,
        severely_late,
        
        -- Payment behavior
        payment_completion,
        payment_timing,
        
        
    FROM payment_behavior
)

SELECT * FROM final