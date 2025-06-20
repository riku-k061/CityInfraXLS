# 🏙️ CityInfraXLS – Urban Infrastructure Maintenance & Analytics System

A standalone Python application for city management teams to track, log, analyze, and report on infrastructure maintenance operations using Excel files as the data backbone. Supports roads, bridges, streetlights, public parks, budgeting, contractor performance, predictive maintenance, complaint resolution, condition scoring, and geospatial enrichment — all through a smart Excel-driven interface.

---

## 📁 Project Structure

```
cityinfraxls/
├── asset_schema.json
├── customer_schema.json
├── severity_matrix.json
├── condition_scoring_schema.json
├── data/
│   ├── assets.xlsx
│   ├── incidents.xlsx
│   ├── tasks.xlsx
│   └── asset_log.xlsx
│
├── utils/
│   ├── incident_handler.py
│   ├── batch_geocoder.py
│   ├── boundary_validator.py
│   ├── geodata_enrichment.py
│   ├── geodata_handler.py
│   └── excel_handler.py
│
├── tests/
│   ├── test_analyze_maintenance.py
│   ├── test_budget_report_generator.py
│   ├── test_delete_maintenance.py
│   ├── test_expense_logger.py
│   ├── test_export_budget_alerts.py
│   ├── test_excel_handler.py
│   ├── test_incident_handler.py
│   ├── test_incidents.py
│   ├── test_tasks.py
│   └── test_assets.py
│
├── asset_task.py
├── register_task.py
├── delete_task.py
├── register_asset.py
├── query_assets.py
├── delete_asset.py
├── register_incident.py
├── query_incidents.py
├── delete_incident.py
├── requirements.txt          # Project dependencies
├── run_tests.py           # Run all of tests
└── README.md
```

---

## ⚙️ How to Run

1. **Install dependencies**:

```bash
pip install -r requirements.txt
```

2. **Run specific modules**:

```bash
python <filename> <param>
```

---

## 🔍 Key Highlights

| # | Module                             | Status      | Key Features                                                                                |
| - | ---------------------------------- | ----------- | ------------------------------------------------------------------------------------------- |
| 1 | Asset Registry                     | ✅ Completed | CRUD via `assets.xlsx`, type-specific validation, filtering, and grouped exports            |
| 2 | Incident & Damage Reporting        | ✅ Completed | Severity-based SLA matrix, asset linkage validation, status tracking, SLA violation export  |
| 3 | Contractor Management & Tasks      | ✅ Completed | Region/type matching, task assignment, overdue alerts, performance metrics export           |
| 4 | Maintenance History & Forecasting  | ✅ Completed | Historical log, high-frequency flags, lifecycle threshold warnings, 3-month schedule export |
| 5 | Complaint Management Workflow      | ✅ Completed | Category routing, SLA tracking, satisfaction scores, monthly trend reports                  |
| 6 | Budget Allocation & Spend Tracking | ✅ Completed | Department & asset budgeting, expense linkage, overrun detection, department-wise summaries |
| 7 | Infrastructure Condition Scoring   | ✅ Completed | Condition score calculation, regional risk ranking, top-10 critical zones export            |
| 8 | Geospatial Enrichment Service      | ✅ Completed | Geocoding missing locations, boundary validation, enriched `geodata.xlsx`, success summary  |

---

## 🧪 Unit Test Results
⚙️ How to Run

Run all of tests using [run_tests.py](https://github.com/riku-k061/CityInfraXLS/blob/main/run_tests.py) 

```bash
python run_tests.py 
```

📸 Screenshots:

| Case          | Link                                                                                          |
| -------------------- | --------------------------------------------------------------------------------------------- |
| ✅ Unit test result 1 | [View](https://drive.google.com/file/d/1iUv_Qr6YRFV0UYQbPnWoygdiA0mUizPv/view?usp=drive_link) |
| ✅ Unit test result 2 | [View](https://drive.google.com/file/d/1SKLOOqkrsjqduPJ8gsVndPDhf8R0-uMk/view?usp=drive_link) |
| ✅ Unit test result 1 | [View](https://drive.google.com/file/d/1AYAfl6i4gltsRrRg6KT4P3ns6mYCPwIk/view?usp=drive_link) |

📸 Test Coverage & Results Screenshots:

| Description                | Link                                                                                                  |
| -------------------------- | ----------------------------------------------------------------------------------------------------- |
| ✅ Asset Registry Tests     | [View](https://drive.google.com/file/d/1U3Tkvvyln48pwDkQyQ38SXSIEyduuLUF/view?usp=drive_link)       |
| ✅ Incident Reporting Tests | [View](https://drive.google.com/file/d/1-NB98TW4IMaERy4JIvG12xZMQmcN-nvG/view?usp=drive_link)       |
| ✅ Contractor Module Tests  | [View](https://drive.google.com/file/d/1uRW48uI7EqZGj-5ZDm0Wnjh1rsWJP-69/view?usp=drive_link)       |
| ✅ Maintenance Module Tests | [View](https://drive.google.com/file/d/1RJzKQDKQ554WNlt6w2EjyX37l_T1VcAI/view?usp=drive_link)       |
| ✅ Complaint Module Tests   | [View](https://drive.google.com/file/d/1LJno_h46pt9owl-mEvKQyth762U4yvcW/view?usp=drive_link)       |
| ✅ Budgeting Module Tests   | [View](https://drive.google.com/file/d/1dCH_tPTuEKVpSRdOIwglt4CuYK7R81dc/view?usp=drive_link)       |
| ✅ Scoring Module Tests     | [View](https://drive.google.com/file/d/1siEQZfXANfZCfTXXe3OWhlJgB71am61E/view?usp=drive_link)       |
| ✅ Geospatial Module Tests  | [View](https://drive.google.com/file/d/1uYqBLyKVUWcFr_69NzC_eWWcDYyi-yQM/view?usp=drive_link)       |


---

## 🚀 Code Execution Screenshots (Per Conversation)

| Conversation | Description                               | Link                                                                                          |
|--------------|-------------------------------------------|-----------------------------------------------------------------------------------------------|
| 1            | Asset Registry Initialization             | [View](https://drive.google.com/file/d/19B2yPybUS2J1m-UslhBliDIrOIr0nDBX/view?usp=drive_link) |
| 2            | Incident Reporting & SLA Calculation      | [View](https://drive.google.com/file/d/1zNlkf0a1oEDyjO7q_ICqq2GmIjy3f1Mp/view?usp=drive_link) |
| 3            | Contractor Assignment & Task Creation     | [View](https://drive.google.com/file/d/1L7qcUITsO4wRvi-Sy4gOzZlcNdROkqYG/view?usp=drive_link) |
| 4            | Maintenance History Tracking & Forecast   | [View](https://drive.google.com/file/d/1WeQpGfboTEefU0hvUkKXhObaimyWtjpL/view?usp=drive_link) |
| 5            | Public Complaint Workflow & Analytics     | [View](https://drive.google.com/file/d/1mW3B44SL9TVLladfKRHgp-8gMQ0wTi0f/view?usp=drive_link) |
| 6            | Budget Allocation & Overrun Detection     | [View](https://drive.google.com/file/d/1Y-DQzO3AhvXZwaOjJniFW-WAp6NWTYJ7/view?usp=drive_link) |
| 7            | Condition Scoring & Risk Map Export       | [View](https://drive.google.com/file/d/1siEQZfXANfZCfTXXe3OWhlJgB71am61E/view?usp=drive_link) |
| 8            | Geospatial Enrichment & Geodata Summary   | [View](https://drive.google.com/file/d/1ETRS_D7b5CbsoFdcd7irGd32lQJr6SVS/view?usp=drive_link) |

---

## 📦 Dependencies

See [`requirements.txt`](./requirements.txt) for the full list.

