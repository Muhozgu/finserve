{{
    config(
        materialized='table',
        schema='GOLD',
        tags=['gold'],
        description='Risk segmentation: loan counts, exposure, and actual observed default rates by risk category and credit score band, for model calibration checks'
    )
}}

WITH loans AS (
    SELECT * FROM {{ ref('silver_loans') }}
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

-- Latest credit score per customer, joined onto their loans
latest_credit_history AS (
    SELECT *
    FROM {{ ref('silver_credit_history') }}
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY customer_id
        ORDER BY record_date DESC NULLS LAST
    ) = 1
),

loan_risk AS (
    SELECT
        l.loan_id,
        l.customer_id,
        l.loan_amount,
        l.loan_status,
        r.risk_category,
        r.risk_score_band,
        r.probability_of_default,
        c.credit_score_band
    FROM loans l
    LEFT JOIN risk_at_origination r ON l.loan_id = r.loan_id
    LEFT JOIN latest_credit_history c ON l.customer_id = c.customer_id
),

-- Segment by risk_category (model's own risk bucket)
by_risk_category AS (
    SELECT
        'risk_category' AS segment_type,
        COALESCE(risk_category, 'Unknown') AS segment_value,
        COUNT(*) AS total_loans,
        SUM(loan_amount) AS total_exposure,
        AVG(probability_of_default) AS avg_predicted_pd,
        SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) AS actual_defaults,
        ROUND(
            SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) / NULLIF(COUNT(*), 0), 4
        ) AS actual_default_rate
    FROM loan_risk
    GROUP BY COALESCE(risk_category, 'Unknown')
),

-- Segment by credit_score_band (independent view from credit bureau data)
by_credit_score_band AS (
    SELECT
        'credit_score_band' AS segment_type,
        COALESCE(credit_score_band, 'Unknown') AS segment_value,
        COUNT(*) AS total_loans,
        SUM(loan_amount) AS total_exposure,
        AVG(probability_of_default) AS avg_predicted_pd,
        SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) AS actual_defaults,
        ROUND(
            SUM(CASE WHEN loan_status = 'Defaulted' THEN 1 ELSE 0 END) / NULLIF(COUNT(*), 0), 4
        ) AS actual_default_rate
    FROM loan_risk
    GROUP BY COALESCE(credit_score_band, 'Unknown')
),

combined AS (
    SELECT * FROM by_risk_category
    UNION ALL
    SELECT * FROM by_credit_score_band
),

final AS (
    SELECT
        segment_type,
        segment_value,
        total_loans,
        total_exposure,
        avg_predicted_pd,
        actual_defaults,
        actual_default_rate,
        -- Flag: is the model's predicted PD wildly off from actual default rate?
        CASE
            WHEN avg_predicted_pd IS NOT NULL
                AND actual_default_rate IS NOT NULL
                AND ABS(avg_predicted_pd - actual_default_rate) > 0.10
            THEN TRUE
            ELSE FALSE
        END AS calibration_flag
    FROM combined
)

SELECT * FROM final
ORDER BY segment_type, segment_value