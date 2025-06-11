# Internwork
This repository contains code excerpts from projects completed during my internship, developed in collaboration with teammates. The uploaded files reflect only the portions of work I was individually responsible for. Thus, the code provided here is not standalone and may not run independently without the full project context.

## Overview
### 1.[utility.py](./Yield Curve/utility.py)
This module provides mathematical functions to construct, interpolate, and optimize fixed-income yield curves and to smooth issuer credit ratings based on market data. 

Key Skills:
- Python Libraries: `numpy`, `pandas`, `scipy.optimize`, `openpyxl`
- Math & Finance: Hermite polynomial interpolation, Yield-to-Maturity (YTM), credit spreads, bond curve modeling

### 2. [yieldcurve.py](./yieldcurve.py)

This script is used to generate and update the yield curve for bonds with different ratings (AA+ to A-). The outputs are plots of daily yield curves and Excel files storing updated curve data. It automatically:

- Retrieves recent market trading data
- Filters and ranks bond transactions
- Identifies key tenor points
- Optimizes and interpolates yield values
- Visualizes and saves the curve for the current day

Key Skills:
- Python Libraries: `pandas`, `numpy`, `matplotlib`, `scipy.optimize`, `datetime`, `os`
- SQL: extracting and joining multi-source bond market data from a financial database, including `JOIN`, `GROUP BY`, `ORDER BY`, `LIMIT`, and `IS NULL` conditions.
  - Retrieves the last two days of market transactions from multiple tables using SQL joins and subqueries.
  - Filters instruments by issuer type, bond structure, and rating logic (e.g., AA+).
  - Extract benchmark yield curves from a historical database for interpolation and curve construction.

### 3. [timer.py](./timer.py)

This script automatically runs `yieldcurve.py` once every 10 minutes during a specific time window.

Key Skills:
- Python Libraries: `datetime`, `threading`, `os`
