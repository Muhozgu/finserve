from connections.postgres import get_postgres_connection
from connections.snowflake import get_snowflake_connection


def test_postgres():
    connection = None
    try:
        connection = get_postgres_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT version();")
        result = cursor.fetchone()
        print("PostgreSQL connection successful!")
        print(result)
        cursor.close()
    except Exception as e:
        print("PostgreSQL connection failed:")
        print(e)
    finally:
        if connection:
            connection.close()


def test_snowflake():
    connection = None
    try:
        connection = get_snowflake_connection()
        cursor = connection.cursor()
        cursor.execute("""
            SELECT
                CURRENT_USER(),
                CURRENT_DATABASE(),
                CURRENT_SCHEMA(),
                CURRENT_WAREHOUSE();
        """)
        result = cursor.fetchone()
        print("Snowflake connection successful!")
        print(result)
        cursor.close()
    except Exception as e:
        print("Snowflake connection failed:")
        print(e)
    finally:
        if connection:
            connection.close()


if __name__ == "__main__":
    test_postgres()
    test_snowflake()