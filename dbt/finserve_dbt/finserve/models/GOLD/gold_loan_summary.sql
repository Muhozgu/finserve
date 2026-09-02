{{
    config(
        materialized='table',
        schema='GOLD',
        tags=['gold'],
        description='Loan fact table: one row per loan combining application context, customer info, risk assessment, and payment performance'
    )
}}

WITH loans AS (
    SELECT * FROM {{ ref('silver_loans') }}
),

applications AS (
    SELECT * FROM {{ ref('silver_applications') }}
),

customers AS (
    SELECT * FROM {{ ref('silver_customers') }}
),

-- Risk assessment closest to (at or before) loan start date, per loan
risk_at_origination AS (
    SELECT *
    FROM {{ ref('silver_risk_assessment') }}
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY loan_id
        ORDER BY assessment_date DESC NULLS LAST
    ) = 1
),

-- Payment behavior aggregated per loan
payment_agg AS (
    SELECT
        loan_id,
        COUNT(*) AS payments_made,
        SUM(CASE WHEN was_late THEN 1 ELSE 0 END) AS late_payments,
        SUM(CASE WHEN severely_late THEN 1 ELSE 0 END) AS severely_late_payments,
        MAX(days_late) AS max_days_late,
        ROUND(AVG(days_late), 1) AS avg_days_late,
        SUM(amount_due) AS total_amount_due,
        SUM(amount_paid) AS total_amount_paid,
        ROUND(
            SUM(CASE WHEN was_late THEN 0 ELSE 1 END) / NULLIF(COUNT(*), 0), 4
        ) AS on_time_payment_rate,
        MAX(payment_date) AS last_payment_date
    FROM {{ ref('silver_payments') }}
    GROUP BY loan_id
),

final AS (
    SELECT
        -- IDs
        l.loan_id,
        l.application_id,
        l.customer_id,

        -- Customer context
        c.full_name AS customer_name,
        c.country,
        c.employment_status,
        c.annual_income,

        -- Application context
        a.loan_purpose,
        a.application_date,
        a.decision_date,
        a.processing_days,
        a.application_channel,

        -- Loan terms
        l.loan_amount,
        l.interest_rate,
        l.interest_rate_band,
        l.term_months,
        l.monthly_payment,
        l.loan_size_band,
        l.start_date,
        l.maturity_date,
        l.months_remaining,
        l.past_maturity,
        l.loan_status,
        l.estimated_total_cost,
        l.estimated_total_interest,

        -- Risk at origination
        r.risk_score,
        r.risk_score_band,
        r.probability_of_default,
        r.risk_category,
        r.assessment_date AS risk_assessment_date,

        -- Payment performance
        COALESCE(p.payments_made, 0) AS payments_made,
        COALESCE(p.late_payments, 0) AS late_payments,
        COALESCE(p.severely_late_payments, 0) AS severely_late_payments,
        p.max_days_late,
        p.avg_days_late,
        p.total_amount_due,
        p.total_amount_paid,
        p.on_time_payment_rate,
        p.last_payment_date,

        -- Derived flag: is this loan underperforming relative to its risk grade?
        CASE
            WHEN r.risk_category = 'LOW' AND COALESCE(p.late_payments, 0) >= 2 THEN TRUE
            WHEN l.loan_status IN ('Defaulted', 'Delinquent') THEN TRUE
            ELSE FALSE
        END AS is_underperforming

    FROM loans l
    LEFT JOIN applications a ON l.application_id = a.application_id
    LEFT JOIN customers c ON l.customer_id = c.customer_id
    LEFT JOIN risk_at_origination r ON l.loan_id = r.loan_id
    LEFT JOIN payment_agg p ON l.loan_id = p.loan_id
)

SELECT * FROM final