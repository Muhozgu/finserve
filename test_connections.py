# test_connections.py
import os
from dotenv import load_dotenv
import psycopg2
import snowflake.connector

# Load environment variables
load_dotenv()

def test_postgres():
    try:
        print("PostgreSQL connecting ...")
        conn = psycopg2.connect(
            host=os.getenv('POSTGRES_HOST', 'localhost'),
            port=os.getenv('POSTGRES_PORT', '5432'),
            user=os.getenv('POSTGRES_USER'),
            password=os.getenv('POSTGRES_PASSWORD'),
            database=os.getenv('POSTGRES_DATABASE')
        )
        print("PostgreSQL connection successful!")
        conn.close()
        return True
    except Exception as e:
        print(f"PostgreSQL connection failed:\n{e}")
        return False

def test_snowflake():
    try:
        print("Snowflake connecting ...")
        conn = snowflake.connector.connect(
            user=os.getenv('SNOWFLAKE_USER'),
            password=os.getenv('SNOWFLAKE_PASSWORD'),
            account=os.getenv('SNOWFLAKE_ACCOUNT'),
            warehouse=os.getenv('SNOWFLAKE_WAREHOUSE'),
            database=os.getenv('SNOWFLAKE_DATABASE'),
            schema=os.getenv('SNOWFLAKE_SCHEMA')
        )
        print("Snowflake connection successful!")
        conn.close()
        return True
    except Exception as e:
        print(f"Snowflake connection failed:\n{e}")
        return False

if __name__ == "__main__":
    test_postgres()
    test_snowflake()