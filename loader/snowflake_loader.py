#!/usr/bin/env python3
"""
NovaSpin Casino - Snowflake Bronze Loader
==========================================

Loads the three CSVs produced by novaspin_casino_data_generator.py into a
BRONZE schema in Snowflake, creating the database/schema/tables if they
don't already exist. This is the raw-load step of the Bronze -> Silver ->
Gold pipeline described in the project README - all columns are loaded as
strings; typing and cleaning happen in the dbt Silver models, not here.

Required environment variables (see .env.example):
    SNOWFLAKE_ACCOUNT
    SNOWFLAKE_USER
    SNOWFLAKE_PASSWORD
    SNOWFLAKE_ROLE        (optional, default ACCOUNTADMIN)
    SNOWFLAKE_WAREHOUSE   (optional, default NOVASPIN)
    SNOWFLAKE_DATABASE    (optional, default NOVASPIN)

Usage:
    python loader/snowflake_loader.py --data-dir data/raw
"""

import argparse
import os
import sys

import pandas as pd
import snowflake.connector
from snowflake.connector.pandas_tools import write_pandas


# Maps target Bronze table name -> source CSV filename
TABLES = {
    "dim_markets": "dim_markets.csv",
    "dim_players": "dim_players.csv",
    "fact_game_sessions": "fact_game_sessions.csv",
}


def get_connection():
    required = ["SNOWFLAKE_ACCOUNT", "SNOWFLAKE_USER", "SNOWFLAKE_PASSWORD"]
    missing = [v for v in required if not os.environ.get(v)]
    if missing:
        raise EnvironmentError(f"Missing required environment variables: {', '.join(missing)}")

    return snowflake.connector.connect(
        account=os.environ["SNOWFLAKE_ACCOUNT"],
        user=os.environ["SNOWFLAKE_USER"],
        password=os.environ["SNOWFLAKE_PASSWORD"],
        role=os.environ.get("SNOWFLAKE_ROLE", "ACCOUNTADMIN"),
        warehouse=os.environ.get("SNOWFLAKE_WAREHOUSE", "COMPUTE_WH"),
    )


def ensure_database_and_schema(conn, database, schema):
    cur = conn.cursor()
    try:
        cur.execute(f"CREATE DATABASE IF NOT EXISTS {database}")
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {database}.{schema}")
    finally:
        cur.close()


def load_table(conn, database, schema, table_name, csv_path):
    # Loaded as strings on purpose: Bronze is a raw landing zone, typing
    # and cleaning are dbt's job in the Silver layer.
    df = pd.read_csv(csv_path, dtype=str)
    df.columns = [c.upper() for c in df.columns]

    success, n_chunks, n_rows, _ = write_pandas(
        conn,
        df,
        table_name=table_name.upper(),
        database=database,
        schema=schema,
        auto_create_table=True,
        overwrite=True,
    )
    if not success:
        raise RuntimeError(f"Failed to load {table_name} into Snowflake")
    print(f"  -> {table_name}: {n_rows:,} rows loaded ({n_chunks} chunk(s))")


def main():
    parser = argparse.ArgumentParser(description="Load NovaSpin CSVs into a Snowflake BRONZE schema.")
    parser.add_argument("--data-dir", type=str, default="data/raw", help="Directory containing the generated CSVs")
    parser.add_argument("--schema", type=str, default="BRONZE", help="Target Snowflake schema (default: BRONZE)")
    args = parser.parse_args()

    database = os.environ.get("SNOWFLAKE_DATABASE", "NOVASPIN")

    for filename in TABLES.values():
        path = os.path.join(args.data_dir, filename)
        if not os.path.isfile(path):
            print(f"Error: expected file not found: {path}", file=sys.stderr)
            sys.exit(1)

    conn = get_connection()
    try:
        print(f"Connected to Snowflake. Loading into {database}.{args.schema} ...")
        ensure_database_and_schema(conn, database, args.schema)
        for table_name, filename in TABLES.items():
            path = os.path.join(args.data_dir, filename)
            load_table(conn, database, args.schema, table_name, path)
        print("Bronze load complete.")
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)