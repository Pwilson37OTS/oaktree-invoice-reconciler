"""Snowflake access to BULLHORN_MIRROR for the invoice reconciler.

Mirrors the pattern the Tickler portal uses (Management-OTS Apps / Tickler
Machine): one Snowflake service account (role CLAUDEAI, which reads
BULLHORN_MIRROR), key-pair auth, credentials from ``[connections.snowflake]``.

Credentials are resolved in this order so the same code runs in the Streamlit
app (local + Community Cloud), from the CLI, and in tests:
  1. ``st.secrets["connections"]["snowflake"]`` when a Streamlit runtime exists
     (this is how Cloud supplies them — paste secrets.cloud.toml into the box),
  2. ``.streamlit/secrets.toml`` next to this file (local dev),
  3. ``SNOWFLAKE_*`` environment variables (CI / scheduled runs), which win.

Local auth uses ``private_key_file`` (a path); Cloud has no filesystem, so the
PEM goes in secrets as ``private_key_pem`` and is converted to the DER bytes the
connector wants. Both are supported here.

The reconciler opens a fresh connection per reconcile run (runs are infrequent
and the result is cached in session state), which sidesteps the long-lived
session-token expiry the Tickler portal has to defend against.
"""
from __future__ import annotations

import os
import tomllib
from functools import lru_cache
from pathlib import Path

import pandas as pd
import snowflake.connector

# Every query here is written with ? placeholders.
snowflake.connector.paramstyle = "qmark"

_SECRETS = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"

# env var -> connection kwarg
_ENV = {
    "SNOWFLAKE_ACCOUNT": "account",
    "SNOWFLAKE_USER": "user",
    "SNOWFLAKE_PASSWORD": "password",
    "SNOWFLAKE_ROLE": "role",
    "SNOWFLAKE_WAREHOUSE": "warehouse",
    "SNOWFLAKE_DATABASE": "database",
    "SNOWFLAKE_SCHEMA": "schema",
    "SNOWFLAKE_PRIVATE_KEY": "private_key_pem",
}


@lru_cache(maxsize=1)
def _config() -> dict:
    cfg: dict = {}

    # 1. Streamlit secrets (how Cloud supplies credentials).
    try:
        import streamlit as st  # noqa: PLC0415 — optional in CLI/tests
        sf = st.secrets.get("connections", {}).get("snowflake", {})
        if sf:
            cfg.update(dict(sf))
    except Exception:
        pass  # no Streamlit runtime / no secrets configured — fall through

    # 2. Local secrets.toml.
    if not cfg and _SECRETS.exists():
        with _SECRETS.open("rb") as fh:
            cfg.update(tomllib.load(fh).get("connections", {}).get("snowflake", {}))

    # 3. Environment wins, so CI / scheduled runs can override without a file.
    for env_name, key in _ENV.items():
        value = os.environ.get(env_name)
        if value:
            cfg[key] = value

    if not cfg.get("account"):
        raise RuntimeError(
            "No Snowflake credentials. Provide a [connections.snowflake] section "
            "in .streamlit/secrets.toml (local) or the Streamlit Cloud Secrets box, "
            "or set SNOWFLAKE_* environment variables. See .streamlit/secrets.toml.example."
        )
    return cfg


def _connect():
    cfg = dict(_config())

    # Cloud ships the key as a PEM string in secrets; the connector wants DER.
    # Local uses private_key_file (a path), which passes straight through.
    pem = cfg.pop("private_key_pem", None)
    if pem:
        from cryptography.hazmat.primitives import serialization

        raw = pem.encode() if isinstance(pem, str) else pem
        key = serialization.load_pem_private_key(raw, password=None)
        cfg["private_key"] = key.private_bytes(
            encoding=serialization.Encoding.DER,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    cfg.pop("private_key_file_pwd", None)
    return snowflake.connector.connect(**cfg)


def query(sql: str, params: list | None = None) -> pd.DataFrame:
    """Run a read query and return a DataFrame (empty frame if no rows)."""
    with _connect() as conn, conn.cursor() as cur:
        cur.execute(sql, params or [])
        if not cur.description:
            return pd.DataFrame()
        try:
            return cur.fetch_pandas_all() if cur.rowcount else pd.DataFrame(
                columns=[c[0] for c in cur.description]
            )
        except snowflake.connector.errors.NotSupportedError:
            rows = cur.fetchall()
            return pd.DataFrame(rows, columns=[c[0] for c in cur.description])


def sync_freshness() -> dict:
    """When BULLHORN_MIRROR last received a CData write, for the BillableCharge
    table. Used to show data age in the app. Returns {} on any error so a
    freshness probe never breaks a reconcile run."""
    try:
        df = query(
            'SELECT MAX("DateLastModified") AS last_modified, '
            'MAX("PeriodEndDate") AS max_period_end '
            'FROM BULLHORN_MIRROR."RAW"."BillableCharge"'
        )
        if df.empty:
            return {}
        row = df.iloc[0]
        return {
            "last_modified": str(row.get("LAST_MODIFIED") or ""),
            "max_period_end": str(row.get("MAX_PERIOD_END") or ""),
        }
    except Exception:
        return {}
