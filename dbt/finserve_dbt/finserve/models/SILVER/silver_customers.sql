{{
    config(
        materialized='table',
        schema='SILVER',
        tags=['silver'],
        description='Cleaned and standardized customers data'
    )
}}

WITH bronze_customers AS (
    SELECT * FROM {{ ref('bronze_customers') }}
),

-- Clean and standardize gender
gender_cleaned AS (
    SELECT
        *,
        CASE 
            WHEN UPPER(TRIM(gender)) IN ('M', 'MALE', 'MALE ') THEN 'Male'
            WHEN UPPER(TRIM(gender)) IN ('F', 'FEMALE', 'FEMALE ') THEN 'Female'
            WHEN UPPER(TRIM(gender)) IN ('NON-BINARY', 'NON BINARY', 'NONBINARY') THEN 'Non-Binary'
            WHEN gender IS NULL OR TRIM(gender) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(gender))
        END AS gender_standardized,
        -- Flag for valid gender
        CASE 
            WHEN UPPER(TRIM(gender)) IN ('M', 'MALE', 'MALE ', 'F', 'FEMALE', 'FEMALE ',
                                         'NON-BINARY', 'NON BINARY', 'NONBINARY') THEN TRUE
            ELSE FALSE
        END AS has_valid_gender
    FROM bronze_customers
),

-- Clean and standardize employment status
employment_cleaned AS (
    SELECT
        *,
        CASE 
            WHEN UPPER(TRIM(employment_status)) IN ('EMPLOYED', 'EMPLOYED ', 'EMP.', 'EMP', 'EMPLOYE') THEN 'Employed'
            WHEN UPPER(TRIM(employment_status)) IN ('SELF-EMPLOYED', 'SELF EMPLOYED', 'SELFEMPLOYED', 'FREELANCE', 'FREELANCER') THEN 'Self-Employed'
            WHEN UPPER(TRIM(employment_status)) IN ('UNEMPLOYED', 'UNEMPLOYED ', 'UNEMP') THEN 'Unemployed'
            WHEN UPPER(TRIM(employment_status)) IN ('RETIRED', 'RETIRED ', 'RET') THEN 'Retired'
            WHEN UPPER(TRIM(employment_status)) IN ('STUDENT', 'STUDENT ') THEN 'Student'
            WHEN UPPER(TRIM(employment_status)) IN ('FT', 'FULL TIME', 'FULL-TIME', 'FULLTIME') THEN 'Full-Time'
            WHEN employment_status IS NULL OR TRIM(employment_status) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(employment_status))
        END AS employment_status_standardized,
        -- Flag for valid employment status
        CASE 
            WHEN UPPER(TRIM(employment_status)) IN ('EMPLOYED', 'EMPLOYED ', 'EMP.', 'EMP', 'EMPLOYE',
                                                     'SELF-EMPLOYED', 'SELF EMPLOYED', 'SELFEMPLOYED', 'FREELANCE', 'FREELANCER',
                                                     'UNEMPLOYED', 'UNEMPLOYED ', 'UNEMP',
                                                     'RETIRED', 'RETIRED ', 'RET',
                                                     'STUDENT', 'STUDENT ',
                                                     'FT', 'FULL TIME', 'FULL-TIME', 'FULLTIME') THEN TRUE
            ELSE FALSE
        END AS has_valid_employment_status
    FROM gender_cleaned
),

-- Clean numeric fields (annual_income, monthly_income)
numeric_cleaned AS (
    SELECT
        *,
        -- Clean annual income (handle negative values as errors)
        CASE 
            WHEN TRY_TO_NUMBER(annual_income) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(annual_income) < 0 THEN NULL  -- Negative income is invalid
            ELSE TRY_TO_NUMBER(annual_income)
        END AS annual_income_clean,
        -- Clean monthly income
        CASE 
            WHEN TRY_TO_NUMBER(monthly_income) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(monthly_income) < 0 THEN NULL  -- Negative income is invalid
            ELSE TRY_TO_NUMBER(monthly_income)
        END AS monthly_income_clean,
        -- Flag for missing annual income
        CASE 
            WHEN annual_income IS NULL OR TRIM(annual_income) = '' OR TRY_TO_NUMBER(annual_income) < 0 THEN TRUE
            ELSE FALSE
        END AS annual_income_missing,
        -- Flag for missing monthly income
        CASE 
            WHEN monthly_income IS NULL OR TRIM(monthly_income) = '' OR TRY_TO_NUMBER(monthly_income) < 0 THEN TRUE
            ELSE FALSE
        END AS monthly_income_missing,
        -- Calculate derived ratio if both exist
        CASE 
            WHEN annual_income_clean > 0 AND monthly_income_clean > 0 
            THEN ROUND(annual_income_clean / (monthly_income_clean * 12), 2)
            ELSE NULL
        END AS income_ratio_check
    FROM employment_cleaned
),

-- Clean and validate date_of_birth
date_cleaned AS (
    SELECT
        *,
        TRY_TO_DATE(date_of_birth) AS date_of_birth_clean,
        -- Calculate age
        DATEDIFF('year', TRY_TO_DATE(date_of_birth), CURRENT_DATE()) AS age_calculated,
        -- Flag for invalid dates (future dates, very old dates)
        CASE 
            WHEN TRY_TO_DATE(date_of_birth) IS NULL THEN 'Invalid'
            WHEN TRY_TO_DATE(date_of_birth) > CURRENT_DATE() THEN 'Future Date'
            WHEN TRY_TO_DATE(date_of_birth) < '1900-01-01' THEN 'Very Old'
            ELSE 'Valid'
        END AS dob_validity,
        -- Flag for unrealistic ages (too young or too old)
        CASE 
            WHEN TRY_TO_DATE(date_of_birth) IS NULL THEN TRUE
            WHEN DATEDIFF('year', TRY_TO_DATE(date_of_birth), CURRENT_DATE()) < 18 THEN TRUE  -- Under 18
            WHEN DATEDIFF('year', TRY_TO_DATE(date_of_birth), CURRENT_DATE()) > 120 THEN TRUE  -- Over 120
            ELSE FALSE
        END AS age_unrealistic
    FROM numeric_cleaned
),

-- Clean country and city
location_cleaned AS (
    SELECT
        *,
        -- Clean country
        CASE 
            WHEN country IS NULL OR TRIM(country) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(country))
        END AS country_standardized,
        -- Clean city
        CASE 
            WHEN city IS NULL OR TRIM(city) = '' THEN 'Unknown'
            ELSE INITCAP(TRIM(city))
        END AS city_standardized,
        -- Flag for missing location info
        CASE 
            WHEN (country IS NULL OR TRIM(country) = '') AND (city IS NULL OR TRIM(city) = '') THEN TRUE
            ELSE FALSE
        END AS location_missing
    FROM date_cleaned
),

-- Clean employment_length_years
employment_length_cleaned AS (
    SELECT
        *,
        -- Clean employment length (handle negative values)
        CASE 
            WHEN TRY_TO_NUMBER(employment_length_years) IS NULL THEN NULL
            WHEN TRY_TO_NUMBER(employment_length_years) < 0 THEN NULL
            ELSE TRY_TO_NUMBER(employment_length_years)
        END AS employment_length_years_clean,
        -- Flag for unrealistic employment length (more than 50 years)
        CASE 
            WHEN TRY_TO_NUMBER(employment_length_years) > 50 THEN TRUE
            ELSE FALSE
        END AS employment_length_unrealistic
    FROM location_cleaned
),

-- Final selection with all cleaned fields
final AS (
    SELECT
        -- IDs
        customer_id,
        
        -- Name fields (no cleaning needed, but can add standardization if needed)
        INITCAP(TRIM(first_name)) AS first_name,
        INITCAP(TRIM(last_name)) AS last_name,
        CONCAT(INITCAP(TRIM(first_name)), ' ', INITCAP(TRIM(last_name))) AS full_name,
        
        -- Demographics
        date_of_birth_clean AS date_of_birth,
        age_calculated AS age,
        dob_validity,
        age_unrealistic,
        
        -- Gender
        gender_standardized AS gender,
        has_valid_gender,
        
        -- Location
        country_standardized AS country,
        city_standardized AS city,
        location_missing,
        
        -- Employment
        employment_status_standardized AS employment_status,
        employment_length_years_clean AS employment_length_years,
        has_valid_employment_status,
        employment_length_unrealistic,
        
        -- Income
        annual_income_clean AS annual_income,
        monthly_income_clean AS monthly_income,
        annual_income_missing,
        monthly_income_missing,
        income_ratio_check,
        
        -- Metadata
        created_at
        
    FROM employment_length_cleaned
)

SELECT * FROM final