# Portfolio Risk Assessment and Optimization in Excel

**Advanced Excel Financial Modeling · Portfolio Risk Analytics · VaR / ES · Monte Carlo Simulation · Backtesting**

This repository presents an Excel-based financial risk modeling project focused on building, optimizing, and validating a mixed investment portfolio consisting of U.S. equities and Russian government bond yield-curve instruments.

The project is designed to demonstrate **advanced Excel modeling skills** in a finance and risk analytics context. The main artifact is a large interconnected Excel workbook that includes data preparation, portfolio construction, risk measurement, optimization, and model validation.

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Business Problem](#business-problem)
3. [Project Objectives](#project-objectives)
4. [Portfolio Specification](#portfolio-specification)
5. [Data Sources and Preparation](#data-sources-and-preparation)
6. [Methodology](#methodology)
7. [Part I — Portfolio Construction and Optimization](#part-i--portfolio-construction-and-optimization)
8. [Part II — Risk Measures and Backtesting](#part-ii--risk-measures-and-backtesting)
9. [Excel Model Overview](#excel-model-overview)
10. [Key Results](#key-results)
11. [Business Value](#business-value)
12. [Skills Demonstrated](#skills-demonstrated)
13. [Repository Contents](#repository-contents)
14. [Author](#author)

---

## Executive Summary

This project builds a full-scale **portfolio risk assessment and optimization model in Microsoft Excel**.

The model evaluates a mixed portfolio consisting of:

- 10 large U.S. equities;
- 5 Russian government bond maturity nodes;
- USD/RUB exchange rate adjustment;
- RUB-denominated returns;
- daily and monthly investment horizons.

The project covers the complete analytical workflow:

```text
Raw Market Data
→ Data Cleaning and Alignment
→ Return Calculation
→ Portfolio Construction
→ Optimization
→ Risk Measurement
→ VaR / ES Estimation
→ Monte Carlo Simulation
→ Backtesting
→ Final Portfolio Evaluation
```

The key value of the project is not only the financial analysis itself, but also the **advanced Excel implementation**. 
The project demonstrates the ability to build a complex, formula-driven Excel model suitable for financial analytics, risk assessment, and investment decision support.

---

## Business Problem

Investment portfolios are exposed to multiple types of risk:

- equity market volatility;
- exchange rate effects;
- interest rate movements;
- bond duration risk;
- tail risk during extreme market conditions;
- changes in the relationship between asset classes over different time horizons.

A portfolio manager or financial analyst needs a model that can answer practical questions:

- What is the risk of the current portfolio?
- How does risk change between daily and monthly horizons?
- Can portfolio weights be optimized to reduce volatility?
- How large can losses be under adverse market scenarios?
- Does the risk model pass statistical validation?
- How do bonds and equities interact in a mixed 60/40 portfolio?

This project addresses these questions through an Excel-based analytical model.

---

## Project Objectives

The project was designed to build a practical investment risk model with the following objectives:

| Objective | Description |
|---|---|
| Portfolio construction | Build equity, bond, and combined 60/40 portfolios |
| Risk measurement | Estimate volatility, beta, Sharpe ratio, VaR, and Expected Shortfall |
| Portfolio optimization | Find minimum-variance and mean-variance portfolio allocations |
| Fixed-income analytics | Estimate bond prices, duration, Macaulay duration, and convexity |
| Scenario simulation | Use Monte Carlo simulation to estimate potential portfolio outcomes |
| Model validation | Backtest VaR models using exceedance frequency and binomial testing |
| Excel implementation | Build the full workflow in a structured Excel model |

---

## Portfolio Specification

The project analyzes a mixed portfolio denominated in RUB.

| Component | Specification |
|---|---|
| Portfolio structure | 60% equities / 40% bonds |
| Equity universe | 10 U.S. stocks |
| Bond universe | 5 Russian OFZ yield curve maturity nodes |
| Currency | All returns converted to RUB |
| FX conversion | USD/RUB daily exchange rate |
| Period | October 1, 2023 – September 30, 2025 |
| Frequency | Daily and 1-month return horizons |
| Benchmark | S&P 500, converted to RUB |

### Equity Holdings

| Ticker | Company | Sector |
|---|---|---|
| INTC | Intel Corporation | Semiconductors |
| AAPL | Apple Inc. | Technology / Consumer Electronics |
| MSFT | Microsoft Corporation | Software & Cloud Computing |
| CSCO | Cisco Systems | Networking Equipment |
| DIS | The Walt Disney Company | Media & Entertainment |
| FDX | FedEx Corporation | Transportation & Logistics |
| GM | General Motors | Automotive |
| CVX | Chevron Corporation | Energy / Oil & Gas |
| BA | The Boeing Company | Aerospace |
| AMZN | Amazon.com Inc. | E-commerce & Cloud Computing |

### Bond Holdings

The bond component is based on Russian government zero-coupon yield curve nodes.

| Maturity | Interpretation |
|---|---|
| 1 Year | Short-term rate exposure |
| 3 Years | Short-to-medium maturity |
| 5 Years | Medium maturity |
| 10 Years | Long-term maturity |
| 30 Years | Very long-term maturity |

---

## Data Sources and Preparation

The model combines equity prices, bond yield curve data, benchmark returns, and exchange rates.

| Data Type | Source / Logic | Usage |
|---|---|---|
| U.S. equities | Historical price data | Equity return calculation |
| S&P 500 | Historical benchmark data | Beta estimation |
| USD/RUB exchange rate | Daily exchange rate series | Conversion of USD prices into RUB |
| OFZ ZCYC | Bank of Russia zero-coupon yield curve | Bond pricing and risk-free rate |
| Risk-free rate | 1-year OFZ ZCYC node | Excess returns and Sharpe ratio |

### Data Preparation Steps

The Excel model performs several preparation steps:

1. Aligns all datasets by date.
2. Converts U.S. stock prices from USD to RUB.
3. Calculates daily and 1-month returns.
4. Computes excess returns using the RUB risk-free rate.
5. Builds covariance and correlation matrices.
6. Prepares return distributions for VaR, ES, and Monte Carlo simulation.

The workbook uses formulas such as `XLOOKUP`, `IFERROR`, and cross-sheet references to align financial data across instruments and dates.

---

## Methodology

### 1. Return Calculation

Daily returns were calculated as:

```text
Return(t) = Price(t) / Price(t-1) - 1
```

Monthly returns were approximated using a 30-day rolling horizon:

```text
Monthly Return(t) = Price(t) / Price(t-30) - 1
```

Log returns were also calculated for distributional analysis and Monte Carlo simulation:

```text
Log Return(t) = LN(1 + Return(t))
```

---

### 2. Portfolio Performance Metrics

The model calculates:

| Metric | Purpose |
|---|---|
| Expected Return | Measures average portfolio performance |
| Standard Deviation | Measures portfolio volatility |
| Sharpe Ratio | Measures risk-adjusted performance |
| Beta | Measures sensitivity to the market benchmark |
| Covariance Matrix | Captures co-movement between assets |
| Correlation Matrix | Measures strength and direction of asset relationships |

---

### 3. Portfolio Optimization

Portfolio weights were optimized under two configurations.

| Optimization Type | Goal |
|---|---|
| Minimum Variance | Minimize portfolio volatility |
| Mean Variance | Increase expected return while accepting higher risk |

The optimization was implemented in Excel using the Solver / OpenSolver logic, based on covariance matrices and portfolio weight constraints.

---

### 4. Fixed-Income Analytics

For the bond portfolio, the model includes:

- bond price estimation;
- effective duration;
- Macaulay duration;
- convexity;
- yield curve dynamics;
- interest rate sensitivity.

Bond prices were calculated using present value logic based on the yield curve.

---

### 5. Risk Modeling

The project estimates downside risk using:

| Risk Measure | Description |
|---|---|
| VaR 95% | Estimates maximum expected loss under normal market conditions |
| Expected Shortfall 95% | Estimates average loss beyond the VaR threshold |
| Historical Simulation | Uses empirical return distribution |
| Monte Carlo Simulation | Simulates possible portfolio outcomes |
| Backtesting | Validates VaR estimates statistically |

---

## Part I — Portfolio Construction and Optimization

### Equity Portfolio

The equity portfolio was built from 10 U.S. companies across different sectors to ensure diversification.

A key feature of the project is that U.S. equity prices were converted into RUB. This means that portfolio returns reflect both:

- U.S. stock price movements;
- USD/RUB exchange rate fluctuations.

This produced an important analytical insight:

> On the daily horizon, RUB-converted U.S. equities showed a partial hedge effect due to currency movements. On the monthly horizon, this effect weakened and equities became more aligned with the market benchmark.

### Stock Portfolio Results

| Metric | 1-Day Minimum Variance | 1-Month Minimum Variance |
|---|---:|---:|
| Expected Return | 0.02% | 0.99% |
| Standard Deviation | 1.78% | 7.29% |
| Sharpe Ratio | -0.0086 | -0.0237 |
| Beta | -0.0502 | 1.0684 |

The 1-day portfolio had low volatility and slightly negative beta, while the 1-month portfolio showed stronger market exposure.

---

### Bond Portfolio

The bond portfolio was constructed using five OFZ maturity nodes: 1Y, 3Y, 5Y, 10Y, and 30Y.

The model compares daily and monthly bond portfolio behavior.

| Metric | 1-Day Minimum Variance | 1-Month Minimum Variance |
|---|---:|---:|
| Expected Return | 0.02% | -0.42% |
| Standard Deviation | 0.21% | 1.00% |
| Sharpe Ratio | -0.0792 | -0.07 |

The bond portfolio acted as the stabilizing component of the final 60/40 allocation.

---

### Final 60/40 Portfolio

The final portfolio combines optimized equity and bond components.

| Metric | 1-Day Portfolio | 1-Month Portfolio |
|---|---:|---:|
| Optimal Sharpe Ratio | -0.0368 | -0.0622 |
| Optimal Beta | -0.0301 | 0.6411 |
| Bond Effective Duration | 5.34 years | 4.89 years |
| Bond Macaulay Duration | 6.08 years | 5.59 years |
| Bond Convexity | 7.63 | 7.63 |

The final portfolio demonstrates two different risk profiles:

- **1-day horizon:** defensive and close to market-neutral;
- **1-month horizon:** moderately market-aligned with higher equity sensitivity.

---

## Part II — Risk Measures and Backtesting

### VaR and Expected Shortfall

The project estimates Value-at-Risk and Expected Shortfall for both current and optimized portfolios.

VaR answers the question:

> What is the maximum expected loss under normal conditions at a given confidence level?

Expected Shortfall answers a more conservative question:

> If losses exceed VaR, what is the average expected loss?

This distinction is important because ES captures tail risk more effectively than VaR.

---

### Backtesting Logic

The project validates VaR estimates using a binomial backtesting approach.

The null hypothesis:

```text
H0: observed exceedance probability = 5%
```

The model compares the expected number of VaR exceedances with the actual number of exceedances during the backtesting period.

### Backtesting Results

| Portfolio / Method | VaR 95% | Observed Exceedance Frequency | P-value | Result |
|---|---:|---:|---:|---|
| Historical Simulation | 4.35% | ~12 days out of 276 | 0.7474 | H0 not rejected |
| Current Portfolio | 5.77% | ~16 days out of 276 | 0.6159 | H0 not rejected |
| Optimal Portfolio | — | 2.88% | — | Lower tail-risk frequency |

The backtesting results show that the VaR models were statistically acceptable because deviations from the expected 5% exceedance rate were not significant.

---

### Monte Carlo Simulation

The workbook includes Monte Carlo simulation to model potential future portfolio values.

The simulation compares:

- current portfolio;
- optimized portfolio.

The optimized portfolio showed:

- narrower distribution;
- lower probability of large losses;
- more stable outcomes;
- stronger capital preservation profile.

This confirms the effectiveness of the optimization process.

---

## Excel Model Overview

This is the most important part of the project from a portfolio perspective.

The project demonstrates advanced Excel not as a simple reporting tool, but as a full financial modeling environment.

### Workbook Architecture

The workbook consists of **34 interconnected worksheets** and **111,000+ formulas**.

```text
Raw Data
├── Equity price sheets
├── USD/RUB exchange rate
├── S&P 500 benchmark
├── OFZ yield curve data

Calculation Layer
├── Return calculation
├── Excess returns
├── Log returns
├── Bond pricing
├── Duration
├── Convexity

Risk and Optimization Layer
├── Equity portfolio optimization
├── Bond portfolio optimization
├── Joint 60/40 portfolio
├── VaR
├── Expected Shortfall
├── Monte Carlo simulation
├── Backtesting

Output Layer
├── Final portfolio summary
├── Risk metrics
├── Optimized allocations
├── Key tables and charts
```

---

### Suggested Visuals for GitHub README

To make the repository more visual, I recommend adding screenshots from the Excel file.

Create a folder:

```text
images/
```

and add the following screenshots.

---

#### 1. Workbook Structure

```markdown
![Workbook Structure](images/workbook_structure.png)
```

**What to screenshot:**  
A view of the Excel workbook showing many worksheet tabs.

**Why it matters:**  
This immediately shows that the project is not a simple spreadsheet, but a complex multi-sheet financial model.

**Suggested caption:**

> The Excel workbook includes 34 interconnected worksheets covering data preparation, portfolio construction, risk modeling, optimization, and validation.

---

#### 2. Formula-Driven Financial Model

```markdown
![Formula Model](images/formula_model.png)
```

**What to screenshot:**  
A calculation-heavy sheet with formulas visible or a dense calculation table.

**Why it matters:**  
This demonstrates advanced Excel modeling and cross-sheet formula logic.

**Suggested caption:**

> Example of formula-driven calculations used for return estimation, bond pricing, and portfolio risk metrics.

---

#### 3. Portfolio Optimization Output

```markdown
![Portfolio Optimization](images/portfolio_optimization.png)
```

**What to screenshot:**  
The final portfolio allocation table or optimized weights for stocks and bonds.

**Why it matters:**  
This shows the transition from raw data to actionable investment allocation.

**Suggested caption:**

> Optimized portfolio weights for equities, bonds, and the combined 60/40 portfolio.

---

#### 4. Yield Curve Dynamics

```markdown
![Yield Curve](images/yield_curve_dynamics.png)
```

**What to screenshot:**  
The chart showing quarterly OFZ yield curve changes.

**Why it matters:**  
This gives visual evidence of fixed-income analysis and interest-rate risk modeling.

**Suggested caption:**

> Quarterly OFZ zero-coupon yield curve dynamics used for bond pricing and duration analysis.

---

#### 5. Monte Carlo Simulation

```markdown
![Monte Carlo Simulation](images/monte_carlo_simulation.png)
```

**What to screenshot:**  
Histogram or distribution chart for current and optimized portfolio simulations.

**Why it matters:**  
This makes risk modeling easier to understand visually.

**Suggested caption:**

> Monte Carlo simulation comparing the distribution of current and optimized portfolio outcomes.

---

#### 6. Final Risk Summary

```markdown
![Final Risk Summary](images/final_risk_summary.png)
```

**What to screenshot:**  
The final summary table with Sharpe ratios, beta, duration, convexity, VaR, or ES.

**Why it matters:**  
This is the most recruiter-friendly visual: it shows the final analytical output of the model.

**Suggested caption:**

> Final portfolio risk summary combining performance, sensitivity, and downside risk metrics.

---

### Key Excel Techniques Used

| Excel Technique | How It Was Used |
|---|---|
| XLOOKUP | Date alignment across multiple market datasets |
| IF / IFERROR | Handling missing observations and non-trading days |
| PV | Bond pricing from yield curve inputs |
| DURATION | Macaulay duration calculation |
| LN / EXP | Log returns and return transformations |
| NORMINV / RAND | Monte Carlo random scenario generation |
| PERCENTILE | VaR estimation from return distributions |
| BINOM.DIST | VaR backtesting and hypothesis testing |
| SUMPRODUCT | Portfolio return and weighted metric calculation |
| VAR / STDEV | Volatility and variance estimation |
| Solver / OpenSolver | Portfolio weight optimization |

### Formula Scale

| Formula / Function Group | Approximate Usage |
|---|---:|
| Total formulas | 111,000+ |
| XLOOKUP | 8,700+ |
| PV | 10,700+ |
| DURATION | 2,500+ |
| LN | 7,000+ |
| IF / IFERROR | 27,000+ |
| NORMINV / RAND | 3,700+ |
| VAR / STDEV | 200+ |

This level of formula usage demonstrates practical advanced Excel skills, including financial modeling, risk analytics, and model structuring.

---

## Key Results

| Finding | Interpretation |
|---|---|
| Minimum-variance optimization reduced risk | The optimized equity portfolio had lower volatility than the baseline allocation |
| Daily and monthly horizons behaved differently | Daily returns showed FX-driven hedge effects, while monthly returns were more market-aligned |
| Bonds stabilized the final portfolio | The bond component reduced volatility and helped control downside risk |
| VaR models passed backtesting | Exceedance frequencies were not statistically different from the expected 5% level |
| ES provided a more conservative risk view | Expected Shortfall showed that tail losses could exceed VaR estimates |
| Monte Carlo confirmed optimization value | Optimized portfolio outcomes were more concentrated and less exposed to extreme losses |
| Advanced Excel was used as the main analytical platform | The full model was implemented through linked formulas, Solver logic, and financial functions |

---

## Business Value

The project shows how Excel can be used as a practical decision-support tool for financial analysis.

A similar model could support:

- investment portfolio monitoring;
- risk reporting;
- asset allocation decisions;
- treasury analytics;
- financial planning;
- stress testing;
- risk model validation.

The main value of the project is the ability to translate complex financial risk concepts into a structured Excel model that can be inspected, adjusted, and reused.

---

## Skills Demonstrated

| Area | Skills |
|---|---|
| Advanced Excel | XLOOKUP, IFERROR, PV, DURATION, SUMPRODUCT, NORMINV, PERCENTILE, BINOM.DIST, Solver |
| Financial Modeling | Portfolio construction, bond pricing, yield curve modeling, duration, convexity |
| Risk Analytics | VaR, Expected Shortfall, backtesting, Monte Carlo simulation |
| Investment Analysis | Asset allocation, equity and bond portfolio analysis, Sharpe ratio, beta |
| Data Analysis | Data alignment, return calculation, covariance and correlation analysis |
| Business Analytics | Decision support, performance evaluation, risk-return trade-off interpretation |

---

## Repository Contents

```text
README.md
investment_portfolio_risk_model.xlsx
images/
├── workbook_structure.png
├── formula_model.png
├── portfolio_optimization.png
├── yield_curve_dynamics.png
├── monte_carlo_simulation.png
└── final_risk_summary.png
```

The Excel workbook is included as the main project artifact.

The original written report is not included because the repository is designed to present the project as a clean, employer-oriented portfolio case.

---

## Project Positioning

This project is best understood as an **Advanced Excel Financial Modeling and Risk Analytics case**.

It demonstrates the ability to:

- build large structured Excel models;
- work with multi-source financial data;
- implement portfolio optimization logic;
- calculate financial risk metrics;
- validate risk models statistically;
- communicate results in a business-friendly way.

For recruiters and hiring managers, the project mainly signals strong practical Excel skills combined with financial analytics and structured problem-solving.

---

## Author

**Maksim Dumanov**

HSE University — International Business and Economics

Research Areas: Business Analytics • Financial Analytics • Investment Analysis • Risk Modeling • Advanced Excel
