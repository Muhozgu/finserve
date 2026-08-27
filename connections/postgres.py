import os
import psycopg2
from dotenv import load_dotenv

load_dotenv()


def get_postgres_connection():
    """
    Create and return a PostgreSQL database connection.
    """

    connection = psycopg2.connect(
        host=os.getenv("POSTGRES_HOST"),
        port=os.getenv("POSTGRES_PORT"),
        database=os.getenv("POSTGRES_DATABASE"),
        user=os.getenv("POSTGRES_USER"),
        password=os.getenv("POSTGRES_PASSWORD")
    )

    return connection