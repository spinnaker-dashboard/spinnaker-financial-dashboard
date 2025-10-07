🧮 Growth & Delta Calculations

| Column | Computation                       | Description                                |
| :----- | :-------------------------------- | :----------------------------------------- |
| H      | (RawCurrent - RawPrevious) / RawPrevious | Month-over-month growth rate for raw data. |
| L      | (BQCurrent - BQPrevious) / BQPrevious | Month-over-month growth rate for BigQuery data. |
| M      | BQCurrent - RawCurrent            | Value delta between sources.               |
| N      | "Match" if delta = 0, otherwise "Mismatch" | Validation status indicator.               |
