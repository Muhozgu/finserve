"""
utils/helpers.py
-----------------
Shared, side-effect-free helper functions used by every generator module.

Design notes
------------
- We use a single `numpy.random.Generator` (created once in generate_data.py
  from the configured seed) and pass it explicitly into every function.
  This is the modern, recommended way to get reproducible NumPy randomness
  (as opposed to the legacy global `np.random.seed(...)` API), and it avoids
  subtle bugs where two modules accidentally share/mutate global state in an
  order-dependent way.
- All ID generation is centralized here so ID formats stay consistent
  across entities (e.g. CUST0000001, APP0000001, ...).
- Date helpers operate on pandas Timestamps / numpy arrays and are written
  to be vectorized wherever possible, since row-by-row `.apply()` over
  100k-1M rows is a major performance trap in pandas.
"""

from __future__ import annotations

import numpy as np
import pandas as pd


def make_ids(prefix: str, n: int, width: int = 7) -> np.ndarray:
    """Generate n stable, sequential, zero-padded string IDs like 'CUST0000001'."""
    return np.array([f"{prefix}{i:0{width}d}" for i in range(1, n + 1)])


def random_dates(
    rng: np.random.Generator,
    start: str,
    end: str,
    n: int,
) -> pd.DatetimeIndex:
    """Vectorized uniform-random dates (day granularity) between start and end inclusive."""
    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    total_days = (end_ts - start_ts).days
    offsets = rng.integers(0, total_days + 1, size=n)
    return pd.to_datetime(start_ts) + pd.to_timedelta(offsets, unit="D")


def clip(series: pd.Series, lo, hi) -> pd.Series:
    return series.clip(lower=lo, upper=hi)


def logistic(x: np.ndarray) -> np.ndarray:
    """Numerically stable sigmoid."""
    out = np.empty_like(x, dtype=float)
    pos = x >= 0
    out[pos] = 1.0 / (1.0 + np.exp(-x[pos]))
    exp_x = np.exp(x[~pos])
    out[~pos] = exp_x / (1.0 + exp_x)
    return out


def zscore(x: np.ndarray) -> np.ndarray:
    """Standardize an array; guards against zero variance."""
    std = x.std()
    if std == 0:
        return np.zeros_like(x, dtype=float)
    return (x - x.mean()) / std


def weighted_choice(
    rng: np.random.Generator,
    values: list,
    weights: list,
    size: int,
) -> np.ndarray:
    """Thin wrapper around rng.choice with normalized weights, for readability at call sites."""
    p = np.array(weights, dtype=float)
    p = p / p.sum()
    return rng.choice(values, size=size, p=p)


def sample_indices(rng: np.random.Generator, n_pool: int, n_sample: int, replace: bool = True) -> np.ndarray:
    """Sample row-positions (not IDs) from a pool of size n_pool."""
    return rng.integers(0, n_pool, size=n_sample) if replace else rng.choice(n_pool, size=n_sample, replace=False)
