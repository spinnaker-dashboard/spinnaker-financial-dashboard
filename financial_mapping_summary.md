🗺️ Financials Mapping Summary

| Metric                | Source Table (BigQuery) | Column(s)                   | Spreadsheet Cells        | Description                                  |
| :-------------------- | :---------------------- | :-------------------------- | :----------------------- | :------------------------------------------- |
| Gross Written Premium (GPW) | ext_policies            | SUM(gross_premium_written)  | F7–G7 (Raw), J7–K7 (BQ)  | Total premium written before deductions.     |
| Gross Earned Premium (GPE)  | ext_policies            | SUM(gross_premium_earned)   | F8–G8 (Raw), J8–K8 (BQ)  | Premium earned during the period.            |
| Collected Premium     | ext_cash                | SUM(raw_collected_cash)     | F9–G9 (Raw), J9–K9 (BQ)  | Total cash collected for the period.         |
| Policy Count          | ext_policies            | COUNT(\*)                   | F10–G10 (Raw), J10–K10 (BQ) | Number of active or closed policies.         |
| Paid Loss             | ext_claims              | SUM(indemnity_paid_itd)     | F13–G13 (Raw), J13–K13 (BQ) | Total indemnity payments to date.            |
| Paid Expense (ALAE)   | ext_claims              | SUM(alae_paid_itd)          | F14–G14 (Raw), J14–K14 (BQ) | Allocated loss adjustment expenses to date.  |
| Claim Count           | ext_claims              | COUNT(\*)                   | F15–G15 (Raw), J15–K15 (BQ) | Number of claims filed.                      |
