# Enterprise Risk Project

## Project title
Enterprise Risk Assessment and Analytics for Amazon's AI Warehouse Strategy

## Purpose
This repository supports the ALY6130 Signature Assessment report. It brings together the earlier Amazon SWOT analysis and the final enterprise risk management assessment using the 10 lecturer-approved risks.

## Approved risks used in the report
1. Competitor AI Advantage
2. AI Warehouse Strategy Failure
3. Cybersecurity & Data Breach
4. AI Forecasting Failure
5. System Integration Failure
6. Cloud Service Outage
7. Workforce Adoption Failure
8. Cost Overrun & Delayed ROI
9. Regulatory & Privacy Non-Compliance
10. AI Governance & Ethics Failure

## Repository structure
```text
enterprise-risk-project/
├── README.md
├── requirements.txt
├── data/
│   ├── raw/
│   └── processed/
├── notebooks/
│   ├── eda.ipynb
│   ├── qualitative_analysis.ipynb
│   ├── quantitative_analysis.ipynb
│   └── monte_carlo.ipynb
├── report/
│   ├── final_report.docx
│   ├── final_report.pdf
│   └── charts/
└── src/
    └── create_report.py
```

## Notes on data
The assessment uses synthetic and proxy datasets because real Amazon internal risk event data is not publicly available. The synthetic dataset is designed to represent realistic KRI behaviour such as cloud uptime, forecasting error, security incidents, system integration status, employee adoption, budget variance, and compliance issues.

## Main outputs
- `report/final_report.docx`: editable final report
- `report/final_report.pdf`: submission-ready PDF version
- `data/processed/quantitative_risk_summary.csv`: expected loss and exposure summary
- `data/processed/synthetic_ml_dataset.csv`: synthetic dataset for ML-based prediction
- `data/processed/monte_carlo_summary.csv`: simulation summary statistics
- `report/charts/`: visual outputs used for analysis

## How to reproduce
Install the requirements, then run the notebooks or the script in `src/create_report.py`.

```bash
pip install -r requirements.txt
python src/create_report.py
```
