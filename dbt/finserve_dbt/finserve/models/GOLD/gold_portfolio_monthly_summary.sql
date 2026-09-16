{{
    config(
        materialized='table',
        schema='GOLD',
        tags=['gold'],
        description='Monthly portfolio KPIs: loan origination volume, amounts, defaults, and payment performance by month, for trend reporting'
    )
}}

WITH loans AS (
    SELECT * FROM {{ ref('silver_loans') }}
    WHERE start_date IS NOT NULL
),

payments AS (
    SELECT * FROM {{ ref('silver_payments') }}
    WHERE due_date IS NOT NULL
),

-- Loan origination metrics by month
monthly_originations AS (
    SELECT
        DATE_TRUNC('month', start_date) AS report_month,
        COUNT(*) AS loans_originated,
        SUM(loan_amount) AS total_originated_amount,
        AVG(loan_amount) AS avg_loan_amount,
        AVG(interest_rate) AS avg_interest_rate,
        SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) AS defaulted_loans,
        SUM(CASE WHEN loan_status = 'Delinquent' THEN 1 ELSE 0 END) AS delinquent_loans,
        SUM(CASE WHEN loan_status = 'Paid Off' THEN 1 ELSE 0 END) AS paid_off_loans,
        SUM(CASE WHEN loan_status = 'Active' THEN 1 ELSE 0 END) AS active_loans,
        ROUND(
            SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) / NULLIF(COUNT(*), 0), 4
        ) AS default_rate
    FROM loans
    GROUP BY DATE_TRUNC('month', start_date)
),

-- Payment performance by due month
monthly_payment_performance AS (
    SELECT
        DATE_TRUNC('month', due_date) AS report_month,
        COUNT(*) AS payments_due,
        SUM(CASE WHEN was_late THEN 1 ELSE 0 END) AS late_payments,
        SUM(CASE WHEN severely_late THEN 1 ELSE 0 END) AS severely_late_payments,
        SUM(amount_due) AS total_amount_due,
        SUM(amount_paid) AS total_amount_paid,
        ROUND(
            SUM(amount_paid) / NULLIF(SUM(amount_due), 0), 4
        ) AS collection_rate,
        ROUND(
            SUM(CASE WHEN was_late THEN 0 ELSE 1 END) / NULLIF(COUNT(*), 0), 4
        ) AS on_time_rate
    FROM payments
    GROUP BY DATE_TRUNC('month', due_date)
),

final AS (
    SELECT
        COALESCE(o.report_month, p.report_month) AS report_month,

        -- Origination metrics
        o.loans_originated,
        o.total_originated_amount,
        o.avg_loan_amount,
        o.avg_interest_rate,
        o.defaulted_loans,
        o.delinquent_loans,
        o.paid_off_loans,
        o.active_loans,
        o.default_rate,

        -- Payment metrics
        p.payments_due,
        p.late_payments,
        p.severely_late_payments,
        p.total_amount_due,
        p.total_amount_paid,
        p.collection_rate,
        p.on_time_rate

    FROM monthly_originations o
    FULL OUTER JOIN monthly_payment_performance p
        ON o.report_month = p.report_month
)

SELECT * FROM final
ORDER BY report_month