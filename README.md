# Delta-Hedging an Up-and-Out Call

This derivatives project studies the replication of an up-and-out call option using dynamic and static hedging strategies under the Black-Scholes framework.

It focuses on the instability of barrier option delta near the knock-out level and analyzes how discrete rebalancing generates large hedging errors, asymmetric cost distributions and extreme tail risk. Additionally, the project compares dynamic delta-hedging against the static replication framework of Derman, Ergener & Kani (1995) across different moneyness and maturity scenarios.

## Analysis Structure

The analysis is organized around three sections:

- **Dynamic delta-hedging analysis:** Monte Carlo simulation of discrete delta-hedging strategies, computation of hedging cost distributions, study of skewness, kurtosis, Value-at-Risk (VaR) and Conditional Value-at-Risk (CVaR), and decomposition of hedging costs between knocked-out and surviving trajectories.

- **Sensitivity analysis:** study of hedging instability across moneyness and maturity scenarios, analysis of near-barrier delta dynamics, investigation of Hull Ratios, replication costs and tail losses under different parameter configurations.

- **Static replication and strategy comparison:** implementation of the Derman-Ergener-Kani static replication portfolio, convergence analysis with increasing replication granularity, comparison of dynamic versus static hedging errors, and identification of the conditions under which static replication dominates dynamic hedging.



## Documents

This project is composed of three parts:
- [`spreadsheet.xlsm`](https://github.com/VitalityMigo/YOUR-REPO/blob/main/spreadsheet.xlsm): Excel/VBA implementation used to simulate Monte Carlo trajectories, compute dynamic hedging costs and construct static replication portfolios.
- [`analysis.pdf`](https://github.com/VitalityMigo/YOUR-REPO/blob/main/analysis.pdf): written report presenting the methodology, results and interpretation of the hedging analysis.
- [`vba/`](https://github.com/VitalityMigo/YOUR-REPO/tree/main/vba): exported VBA source code used for pricing, Monte Carlo simulation, delta computation and hedging strategy implementation.

*Note: the framework can be adapted to other barrier option structures and replication frequencies.*
