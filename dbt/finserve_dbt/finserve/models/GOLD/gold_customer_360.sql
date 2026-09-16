{{
    config(
        materialized='table',
        schema='GOLD',
        tags=['gold'],
        description='Customer 360: one row per customer combining demographics, latest credit profile, latest risk assessment, and loan/payment rollups'
    )
}}

WITH customers AS (
    SELECT * FROM {{ ref('silver_customers') }}
),

-- Take the most recent credit history record per customer
latest_credit_history AS (
    SELECT *
    FROM {{ ref('silver_credit_history') }}
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY customer_id
        ORDER BY record_date DESC NULLS LAST
    ) = 1
),

-- Take the most recent risk assessment per customer (across all their loans)
latest_risk_assessment AS (
    SELECT *
    FROM {{ ref('silver_risk_assessment') }}
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY customer_id
        ORDER BY assessment_date DESC NULLS LAST
    ) = 1
),

-- Loan-level rollups per customer
loan_rollups AS (
    SELECT
        customer_id,
        COUNT(*) AS total_loans,
        SUM(CASE WHEN loan_status = 'Active' THEN 1 ELSE 0 END) AS active_loans,
        SUM(CASE WHEN loan_status = 'Paid Off' THEN 1 ELSE 0 END) AS paid_off_loans,
        SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) AS defaulted_loans,
        SUM(CASE WHEN loan_status = 'Delinquent' THEN 1 ELSE 0 END) AS delinquent_loans,
        SUM(loan_amount) AS total_loan_amount,
        AVG(interest_rate) AS avg_interest_rate,
        SUM(CASE WHEN loan_status = 'Active' THEN loan_amount ELSE 0 END) AS outstanding_loan_amount
    FROM {{ ref('silver_loans') }}
    GROUP BY customer_id
),

-- Payment-level rollups per customer
payment_rollups AS (
    SELECT
        customer_id,
        COUNT(*) AS total_payments,
        SUM(CASE WHEN was_late THEN 1 ELSE 0 END) AS late_payments,
        SUM(CASE WHEN severely_late THEN 1 ELSE 0 END) AS severely_late_payments,
        ROUND(
            SUM(CASE WHEN was_late THEN 0 ELSE 1 END) / NULLIF(COUNT(*), 0), 4
        ) AS on_time_payment_rate,
        SUM(amount_due) AS total_amount_due,
        SUM(amount_paid) AS total_amount_paid
    FROM {{ ref('silver_payments') }}
    GROUP BY customer_id
),

final AS (
    SELECT
        -- Customer identity & demographics
        c.customer_id,
        c.full_name,
        c.age,
        c.gender,
        c.country,
        c.city,
        c.employment_status,
        c.employment_length_years,
        c.annual_income,
        c.monthly_income,

        -- Latest credit profile
        lch.credit_score,
        lch.credit_score_band,
        lch.total_outstanding_debt,
        lch.credit_utilization,
        lch.utilization_band,
        lch.total_accounts,
        lch.previous_defaults,
        lch.multiple_defaults,
        lch.chronic_late_payer,

        -- Latest risk assessment
        lra.risk_score,
        lra.risk_score_band,
        lra.probability_of_default,
        lra.probability_band,
        lra.risk_category,
        lra.assessment_date AS latest_assessment_date,

        -- Loan rollups
        COALESCE(lr.total_loans, 0) AS total_loans,
        COALESCE(lr.active_loans, 0) AS active_loans,
        COALESCE(lr.paid_off_loans, 0) AS paid_off_loans,
        COALESCE(lr.defaulted_loans, 0) AS defaulted_loans,
        COALESCE(lr.delinquent_loans, 0) AS delinquent_loans,
        lr.total_loan_amount,
        lr.avg_interest_rate,
        lr.outstanding_loan_amount,

        -- Payment rollups
        COALESCE(pr.total_payments, 0) AS total_payments,
        COALESCE(pr.late_payments, 0) AS late_payments,
        COALESCE(pr.severely_late_payments, 0) AS severely_late_payments,
        pr.on_time_payment_rate,
        pr.total_amount_due,
        pr.total_amount_paid,

        -- Simple customer risk flag for quick filtering
        CASE
            WHEN lra.risk_category = 'HIGH' OR COALESCE(lr.defaulted_loans, 0) > 0 THEN TRUE
            ELSE FALSE
        END AS is_high_risk_customer

    FROM customers c
    LEFT JOIN latest_credit_history lch ON c.customer_id = lch.customer_id
    LEFT JOIN latest_risk_assessment lra ON c.customer_id = lra.customer_id
    LEFT JOIN loan_rollups lr ON c.customer_id = lr.customer_id
    LEFT JOIN payment_rollups pr ON c.customer_id = pr.customer_id
)

SELECT * FROM final