{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized risk assessments data'
    )
}}

WITH bronze_risk_assessments AS (
    SELECT * FROM {{ ref('bronze_risk_assessment') }}
),

-- Clean and standardize risk category
category_cleaned AS (
    SELECT
        *,
        -- Standardize risk category
        CASE 
            WHEN UPPER(TRIM(risk_category)) IN ('LOW', 'LOW ') THEN 'LOW'
            WHEN UPPER(TRIM(risk_category)) IN ('MEDIUM', 'MEDIUM ', 'MED') THEN 'MEDIUM'
            WHEN UPPER(TRIM(risk_category)) IN ('HIGH', 'HIGH ') THEN 'HIGH'
            WHEN risk_category IS NULL OR TRIM(risk_category) = '' THEN 'Unknown'
            ELSE UPPER(TRIM(risk_category))
        END AS risk_category_standardized,
        -- Flag for valid category
        CASE 
            WHEN UPPER(TRIM(risk_category)) IN ('LOW', 'LOW ', 'MEDIUM', 'MEDIUM ', 'MED', 'HIGH', 'HIGH ') THEN TRUE
            ELSE FALSE
        END AS has_valid_category
    FROM bronze_risk_assessments
),

-- Validate numeric fields
numeric_cleaned AS (
    SELECT
        *,
        -- Validate risk_score (should be between 300-850 typically)
        CASE 
            WHEN TRY_TO_NUMBER(risk_score) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(risk_score) < 300 OR TRY_TO_NUMBER(risk_score) > 850 THEN NULL
            ELSE TRY_TO_NUMBER(risk_score)
        END AS risk_score_clean,
        -- Flag for out-of-range risk score
        CASE 
            WHEN TRY_TO_NUMBER(risk_score) IS NOT NULL 
                AND (TRY_TO_NUMBER(risk_score) < 300 OR TRY_TO_NUMBER(risk_score) > 850) THEN TRUE
            ELSE FALSE
        END AS risk_score_out_of_range,
        
        -- Validate probability_of_default (should be between 0-1)
        CASE 
            WHEN TRY_TO_NUMBER(probability_of_default) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(probability_of_default) < 0 OR TRY_TO_NUMBER(probability_of_default) > 1 THEN NULL
            ELSE TRY_TO_NUMBER(probability_of_default)
        END AS probability_of_default_clean,
        -- Flag for out-of-range probability
        CASE 
            WHEN TRY_TO_NUMBER(probability_of_default) IS NOT NULL 
                AND (TRY_TO_NUMBER(probability_of_default) < 0 OR TRY_TO_NUMBER(probability_of_default) > 1) THEN TRUE
            ELSE FALSE
        END AS probability_out_of_range
    FROM category_cleaned
),

-- Validate dates
date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(assessment_date) AS assessment_date_clean,
        -- Flag for future dates
        CASE 
            WHEN TRY_TO_DATE(assessment_date) > CURRENT_DATE() THEN TRUE
            ELSE FALSE
        END AS assessment_date_future,
        -- Flag for very old dates
        CASE 
            WHEN TRY_TO_DATE(assessment_date) < '2020-01-01' THEN TRUE
            ELSE FALSE
        END AS assessment_date_old
    FROM numeric_cleaned
),

-- Create risk score bands for analysis
risk_bands AS (
    SELECT
        *,
        CASE 
            WHEN risk_score_clean >= 750 THEN 'Excellent'
            WHEN risk_score_clean >= 700 THEN 'Good'
            WHEN risk_score_clean >= 650 THEN 'Fair'
            WHEN risk_score_clean >= 600 THEN 'Poor'
            WHEN risk_score_clean IS NOT NULL THEN 'Very Poor'
            ELSE 'Unknown'
        END AS risk_score_band,
        -- Probability of default bands
        CASE 
            WHEN probability_of_default_clean <= 0.05 THEN 'Very Low'
            WHEN probability_of_default_clean <= 0.15 THEN 'Low'
            WHEN probability_of_default_clean <= 0.30 THEN 'Moderate'
            WHEN probability_of_default_clean <= 0.50 THEN 'High'
            WHEN probability_of_default_clean IS NOT NULL THEN 'Very High'
            ELSE 'Unknown'
        END AS probability_band
    FROM date_cleaned
),

-- Check consistency between risk_score and risk_category
consistency_check AS (
    SELECT
        *,
        -- Check if risk_score aligns with risk_category
        CASE 
            WHEN risk_score_clean >= 700 AND risk_category_standardized = 'LOW' THEN 'Consistent'
            WHEN risk_score_clean BETWEEN 600 AND 699 AND risk_category_standardized = 'MEDIUM' THEN 'Consistent'
            WHEN risk_score_clean < 600 AND risk_category_standardized = 'HIGH' THEN 'Consistent'
            WHEN risk_score_clean IS NOT NULL AND risk_category_standardized != 'Unknown' THEN 'Inconsistent'
            ELSE 'Unknown'
        END AS score_category_consistency,
        -- Check if probability matches risk category
        CASE 
            WHEN probability_of_default_clean <= 0.05 AND risk_category_standardized = 'LOW' THEN 'Consistent'
            WHEN probability_of_default_clean BETWEEN 0.05 AND 0.20 AND risk_category_standardized = 'MEDIUM' THEN 'Consistent'
            WHEN probability_of_default_clean > 0.20 AND risk_category_standardized = 'HIGH' THEN 'Consistent'
            WHEN probability_of_default_clean IS NOT NULL AND risk_category_standardized != 'Unknown' THEN 'Inconsistent'
            ELSE 'Unknown'
        END AS probability_category_consistency
    FROM risk_bands
),

-- Final selection
final AS (
    SELECT
        -- IDs
        assessment_id,
        customer_id,
        loan_id,
        
        -- Assessment date
        assessment_date_clean AS assessment_date,
        assessment_date_future,
        assessment_date_old,
        
        -- Risk score
        risk_score_clean AS risk_score,
        risk_score_band,
        risk_score_out_of_range,
        
        -- Probability of default
        probability_of_default_clean AS probability_of_default,
        probability_band,
        probability_out_of_range,
        
        -- Risk category
        risk_category_standardized AS risk_category,
        has_valid_category,
        
        -- Consistency checks
        score_category_consistency,
        probability_category_consistency,
        
        -- Flag for overall data quality
        CASE 
            WHEN risk_score_clean IS NULL 
                OR probability_of_default_clean IS NULL
                OR risk_category_standardized = 'Unknown'
                OR risk_score_out_of_range = TRUE
                OR probability_out_of_range = TRUE
                OR assessment_date_future = TRUE
            THEN TRUE
            ELSE FALSE
        END AS has_data_quality_issue,
        
        -- Model version
        model_version,
        
        -- Created/updated timestamps (if they exist in the data)
        created_at,
        updated_at
        
    FROM consistency_check
)

SELECT * FROM final