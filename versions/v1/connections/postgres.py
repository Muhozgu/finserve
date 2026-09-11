import os
from dotenv import load_dotenv
import psycopg2

load_dotenv(override=True)  # Load environment variables from .env file, allowing overrides


def get_postgres_connection():
    """
    Create and return a psycopg2 connection to the local Postgres container.
    Reads host/port/db/user/password from environment variables.
    """
    return psycopg2.connect(
        host=os.getenv("POSTGRES_HOST", "localhost"),
        port=os.getenv("POSTGRES_PORT", "5432"),
        dbname=os.getenv("POSTGRES_DATABASE"),
        user=os.getenv("POSTGRES_USER"),
        password=os.getenv("POSTGRES_PASSWORD"),
    )