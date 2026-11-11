# M10 PMLS - FINAL EXAM (ASSIGNMENT 2)
## Productionisation of Machine Learning System: FastAPI Deployment with Excel Integration

---

## Executive Summary & System Overview
This project presents an enterprise-grade productionisation of a **Bank Loan Default Risk Prediction System**. A retail financial institution requires an automated, low-latency, and highly explainable mechanism to predict whether loan applicants are likely to default (`DEFAULTER = 1`) or repay their obligations (`DEFAULTER = 0`). 

The complete deliverable encompasses:
1. **Theoretical MLOps Mastery**: In-depth analytical treatises on Model Deployment Strategies (Q1 - 10 Marks) and Model Monitoring & Logging Systems (Q2 - 10 Marks).
2. **Machine Learning Pipeline**: A 500-tree Random Forest Classifier trained on historical customer records, achieving an 82.86% validation accuracy, an ROC-AUC of 0.8305, and 100.0% accuracy on the provided test dataset.
3. **REST API Microservice**: A production FastAPI backend (`main.py`) exposing synchronous `/predict`, `/predict-batch`, `/test-predictions`, and `/health` endpoints with Pydantic data validation and CORS support.
4. **Interactive Macro-Enabled Excel Interface**: A fully formatted Excel Workbook (`Bank_Loan_Default_Prediction.xlsm`) featuring a native Form Control Submit button wired to cross-platform VBA macros (`vba_macro.bas`) that query the live FastAPI endpoint and display predictions dynamically.

---

## Deliverables Checklist & File Inventory

|                                                                                                                                                               File Name                                                                                                                                                                |                                                 Description                                                 |  Status  |
|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------:|:-----------------------------------------------------------------------------------------------------------:|:--------:|
|  [`Final Exam (Assignment).ipynb`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/Final%20Exam%20%28Assignment%29.ipynb)  | Complete Jupyter Notebook addressing Q1, Q2, and all steps of Q3 with pre-executed outputs and benchmarks   | Complete |
|                            [`main.py`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/main.py)                            |             Production FastAPI REST API source code with Swagger documentation and web test GUI             | Complete |
|                       [`model.joblib`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/model.joblib)                       |                  Serialized 500-Tree Random Forest model artifact (Scikit-Learn / Joblib)                   | Complete |
| [`Bank_Loan_Default_Prediction.xlsm`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/Bank_Loan_Default_Prediction.xlsm)   |      Macro-enabled Excel workbook with functional Submit button, single-loan form, and test set sheets      | Complete |
|                      [`vba_macro.bas`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/vba_macro.bas)                      |             Exported cross-platform VBA macro module (supports macOS `curl` & Windows `MSXML2`)             | Complete |
|                     [`BANK LOAN.csv`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/BANK%20LOAN.csv)                     |                      Historical training dataset containing 700 customer loan records                       | Complete |
|                [`BANK LOAN_TEST.csv`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/BANK%20LOAN_TEST.csv)                |                             Evaluation test set containing 10 customer records                              | Complete |
|                   [`requirements.txt`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/requirements.txt)                   |                 Python dependencies required to run the API, notebook, and automation suite                 | Complete |
|                        [`render.yaml`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/render.yaml)                        |                  Infrastructure-as-code cloud deployment manifest for one-click deployment                  | Complete |

---

## Q1. Model Deployment (10 Marks)

### Comprehensive Synthesis
Model deployment transitions an algorithmic model from training code into operational software that services end users, APIs, or automated business workflows. Choosing a deployment strategy depends on latency requirements, data velocity, compute budgets, and failure tolerance:

```
+---------------------------------------------------------------------------------------------------+
|                                     MODEL DEPLOYMENT PARADIGMS                                    |
+-------------------+--------------------+------------------------+---------------------------------+
| Method            | Latency SLA        | Compute Architecture   | Industry Example                |
+-------------------+--------------------+------------------------+---------------------------------+
| 1. Batch Offline  | Hours to Days      | Scheduled Spark / SQL  | Overnight credit risk re-scoring|
| 2. Streaming Event| 50 to 500 ms       | Kafka / Event Broker   | Payment card fraud detection    |
| 3. Serverless FaaS| 100 ms to 5 s      | AWS Lambda / GCP Cloud | Document parsing & loan OCR     |
| 4. Microserve     | 10 to 50 ms        | Docker / Kubernetes    | Core multi-model banking engine |
+-------------------+--------------------+------------------------+---------------------------------+
```

1. **Batch Deployment (Offline Inference)**:
   - ***Mechanics***: Data pipelines run on periodic schedules (cron/Airflow) processing accumulated records. Predictions are stored in databases or data lakes for later retrieval.
   - ***Strengths***: Highly parallelized, high throughput, zero user-facing latency pressure, cost-effective off-peak execution.
   - ***Limitations***: Predictions age quickly; incapable of responding to live user interactions.

2. **Streaming & Event-Driven Deployment**:
   - ***Mechanics***: Models act as consumers subscribed to distributed event streams (Apache Kafka, AWS Kinesis). Events trigger immediate scoring and publish result events to downstream topics.
   - ***Strengths***: Decoupled, asynchronous, resilient against bursty traffic without dropping requests.
   - ***Limitations***: Requires distributed stream infrastructure management and complex stateful event handling.

3. **Serverless Deployment (Function-as-a-Service - FaaS)**:
   - ***Mechanics***: Stateless container functions scale automatically from zero to handle incoming requests on platforms like AWS Lambda or Google Cloud Functions.
   - ***Strengths***: Pay-per-millisecond cost efficiency; zero infrastructure maintenance for intermittent workloads.
   - ***Limitations***: Cold-start latencies; strict memory (<10GB) and execution timeout constraints.

4. **Containerised Microservices Orchestration (Docker & Kubernetes)**:
   - ***Mechanics***: Standardized containers orchestrated via Kubernetes enable advanced deployment strategies like Canary routing (shifting 5% traffic to a new version), Blue-Green cutovers, and Horizontal Pod Autoscaling (HPA).
   - ***Strengths***: Maximum resilience, enterprise governance, zero-downtime rolling updates.
   - ***Limitations***: High operational complexity and engineering overhead.

---

## Q3. FastAPI Deployment with Excel Integration (40 Marks)

### 1. Data Schema & Feature Definition
The model utilises historical customer credit records from `BANK LOAN.csv` (700 rows). Per the assignment prompt:
- **Target Variable ($Y$)**: `DEFAULTER` (Binary: 1 = Defaulted, 0 = Non-Defaulter)
- **Independent Variables ($X$)**: All features except `SN` (Serial Number)
  - `AGE`: Age category (1 = <30, 2 = 30-45, 3 = >45)
  - `EMPLOY`: Years with current employer
  - `ADDRESS`: Years at current residential address
  - `DEBTINC`: Debt-to-Income ratio (%)
  - `CREDDEBT`: Credit card debt balance ($k)
  - `OTHDEBT`: Other consumer debt balance ($k)

### 2. Random Forest Model Architecture & Performance
- **Estimators**: Exactly 500 trees (`n_estimators= 500`, `random_state= 42`)
- Evaluation Benchmark (80/20 Stratified Validation):
  - ***Accuracy***: 82.86%
  - ***ROC-AUC***: 0.8305
  - ***Precision (Defaulter)***: 0.70
  - ***Recall (Defaulter)***: 0.62
  - ***F1-Score (Defaulter)***: 0.66

#### Feature Importance:
```
1. DEBTINC  (Debt-to-Income Ratio)     : 24.95% 
2. CREDDEBT (Credit Card Debt)         : 20.98%
3. EMPLOY   (Employment Tenure)        : 19.10%
4. OTHDEBT  (Other Debt Obligations)   : 17.84%
5. ADDRESS  (Address Tenure)           : 13.31%
6. AGE      (Age Bracket Category)     :  3.83%
```

### 3. BankLoan_Test Predictions & Evaluation
The model was evaluated on all 10 customer records from `BANK LOAN_TEST.csv`:

|   SN   | AGE  | EMPLOY | ADDRESS | DEBTINC | CREDDEBT | OTHDEBT | ACTUAL DEFAULT | ESTIMATED PROBABILITY | PRED CLASS | RISK ASSESSMENT             | CORRECT? |
|:------:|:----:|:------:|:-------:|:-------:|:--------:|:-------:|:--------------:|:---------------------:|:----------:|:----------------------------|:--------:|
| **1**  |  3   |   16   |   10    |   9.1   |  11.30   |  5.01   |       1        |      **67.80%**       |     1      | High Risk (Default Likely)  | **YES**  |
| **2**  |  1   |   8    |    5    |  17.0   |   1.46   |  4.00   |       0        |      **31.60%**       |     0      | Moderate Risk (Underwriter) | **YES**  |
| **3**  |  2   |   15   |   14    |   5.5   |   0.86   |  2.17   |       0        |       **0.00%**       |     0      | Low Risk (Repayment Likely) | **YES**  |
| **4**  |  3   |   13   |   13    |   2.9   |   2.56   |  0.82   |       0        |       **8.60%**       |     0      | Low Risk (Repayment Likely) | **YES**  |
| **5**  |  1   |   2    |    1    |  16.5   |   1.79   |  3.06   |       1        |      **78.40%**       |     1      | High Risk (Default Likely)  | **YES**  |
| **6**  |  3   |   5    |    6    |  10.2   |   0.39   |  2.16   |       0        |       **6.60%**       |     0      | Low Risk (Repayment Likely) | **YES**  |
| **7**  |  2   |   18   |    8    |  31.1   |   3.80   |  16.67  |       0        |      **28.60%**       |     0      | Low Risk (Repayment Likely) | **YES**  |
| **8**  |  3   |   12   |   11    |   3.6   |   0.13   |  1.24   |       0        |       **0.40%**       |     0      | Low Risk (Repayment Likely) | **YES**  |
| **9**  |  1   |   3    |    3    |  24.4   |   1.31   |  3.28   |       1        |      **78.00%**       |     1      | High Risk (Default Likely)  | **YES**  |
| **10** |  2   |   0    |   11    |  19.7   |   2.78   |  2.15   |       0        |      **25.80%**       |     0      | Low Risk (Repayment Likely) | **YES**  |

**Summary Evaluation**: 10 out of 10 correct classifications (**100.0% Test Accuracy**). The model cleanly separates defaulters (average predicted probability of **74.73%**) from non-defaulters (average predicted probability of **14.54%**).

## 3.8 Excel Integration & Macro-Enabled Workbook (`Bank_Loan_Default_Prediction.xlsm`)

### Architecture of the Excel-to-FastAPI System
```
+-----------------------------------------------------------------------------------------+
|                  EXCEL MACRO-ENABLED WORKBOOK (.xlsm) FRONTEND                          |
+-----------------------------------------------------------------------------------------+
|  [Single Customer Prediction Sheet]                                                     |
|    - C6: Age Category (1, 2, 3)                                                         |
|    - C7: Employment Tenure (Years)                                                      |
|    - C8: Address Tenure (Years)                                                         |
|    - C9: Debt-to-Income Ratio (%)                                                       |
|    - C10: Credit Card Debt ($k)                                                         |
|    - C11: Other Debt ($k)                                                               |
|                                                                                         |
|    +-----------------------------------------------+                                    |
|    |  [BUTTON]: "Submit to FastAPI (/predict)"     |                                    |
|    +-----------------------+-----------------------+                                    |
|                            | Click Event triggers VBA Sub SubmitLoanPrediction()        |
+----------------------------|------------------------------------------------------------+
                             |
                             v
+-----------------------------------------------------------------------------------------+
|                           VBA MACRO COMMUNICATION LAYER                                 |
+-----------------------------------------------------------------------------------------+
|  #If Mac Then                                                                           |
|      MacScript("do shell script 'curl -s -X POST -H ...'")                              |
|  #Else                                                                                  |
|      Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")                             |
|      xmlHttp.Open "POST", "http://127.0.0.1:8000/predict", False                        |
|      xmlHttp.send jsonPayload                                                           |
|  #End If                                                                                |
+----------------------------|------------------------------------------------------------+
                             |  HTTP POST /predict  (JSON Payload)
                             v
+-----------------------------------------------------------------------------------------+
|                              FASTAPI BACKEND SERVICE                                    |
+-----------------------------------------------------------------------------------------+
|  1. Ingests JSON payload -> Validates against Pydantic CustomerLoanFeatures             |
|  2. Transforms into DataFrame -> Calls model.predict_proba(df)[:, 1]                    |
|  3. Evaluates credit risk tier & decision recommendations                               |
|  4. Returns JSON: {"default_probability": 0.3160, "risk_tier": "Moderate Risk", ...}    |
+----------------------------+------------------------------------------------------------+
                             |
                             |  HTTP 200 OK Response (JSON)
                             v
+-----------------------------------------------------------------------------------------+
|                    EXCEL WORKSHEET CELL & UI UPDATES                                    |
+-----------------------------------------------------------------------------------------+
|  - C14: Predicted Default Probability -> Formatted as Percentage (e.g., 31.60%)         |
|  - C15: Risk Assessment Tier -> "Moderate Risk (Underwriter Review)"                    |
|  - C16: Credit Decision Recommendation -> "Manual Underwriter Review Required"          |
|  - C17: Backend API Status -> "HTTP 200 OK (FastAPI)" (Green)                           |
|  - C18: Timestamp -> Current System Timestamp                                           |
|  - Popup Alert: Instant MessageBox displaying full credit evaluation summary            |
+-----------------------------------------------------------------------------------------+
```

---

## FastAPI REST API Service (`main.py`)

The REST API is built with **FastAPI** to deliver low latency, schema validation, and automatic OpenAPI documentation:

### Available Endpoints
- `POST /predict`: Primary endpoint. Accepts JSON customer attributes and returns predicted probability of default, risk tier, and credit recommendation.
- `POST /predict-batch`: Accepts a list of customer objects and returns batch predictions.
- `GET /test-predictions`: Returns predictions and metrics for the 10 customers from `BANK LOAN_TEST.csv`.
- `GET /health`: Returns service health status and model metadata.
- `GET /interactive-tester`: Browser-based testing interface.
- `GET /docs`: Interactive Swagger UI documentation.

### Example Request (`POST /predict`):
```bash
curl -X POST "http://127.0.0.1:8000/predict" \
     -H "Content-Type: application/json" \
     -d '{"AGE": 1, "EMPLOY": 8, "ADDRESS": 5, "DEBTINC": 17.0, "CREDDEBT": 1.46, "OTHDEBT": 4.00}'
```

### Example Response:
```json
{
  "default_probability": 0.316,
  "default_probability_pct": "31.60%",
  "prediction": 0,
  "prediction_label": "Non-Defaulter",
  "risk_tier": "Moderate Risk",
  "recommendation": "Manual Underwriter Review Required",
  "input_features": {
    "AGE": 1.0,
    "EMPLOY": 8.0,
    "ADDRESS": 5.0,
    "DEBTINC": 17.0,
    "CREDDEBT": 1.46,
    "OTHDEBT": 4.0
  }
}
```

---

## Excel Macro-Enabled Interface (`Bank_Loan_Default_Prediction.xlsm`)

The workbook provides a business-ready interface that bridges desktop banking workflows with the Python backend:

### Sheet Structure
1. **`Single Customer Prediction`**:
   - Customer financial and demographic input cells (`C6:C11`).
   - Form Control Submit to FastAPI (/predict) button tied to macro `SubmitLoanPrediction()`.
   - Dynamic output cells displaying default probability (`C14`), risk tier (`C15`), recommendation (`C16`), API status (`C17`), and timestamp (`C18`).
   - Side-by-side system operating instructions.
2. **`BankLoan_Test Predictions`**:
   - Complete table of all 10 test customers pre-populated with features and predicted default probabilities.
   - Run Batch API Prediction button tied to macro `BatchPredictTestLoans()` that live-queries the API across all 10 rows.
3. **`Model Metrics & Architecture`**:
   - Reference specifications for model configuration, validation benchmarks, and feature importance rankings.

### Cross-Platform VBA Architecture (`vba_macro.bas`)
The embedded VBA code supports **both Windows and macOS** seamlessly through conditional compilation:
```vba
#If Mac Then
    ' macOS execution path using curl
    Dim macCurlCmd As String
    macCurlCmd = "do shell script ""curl -s -X POST -H \""Content-Type: application/json\"" -d '" & jsonPayload & "' """ & apiUrl & """"""
    responseText = MacScript(macCurlCmd)
#Else
    ' Windows execution path using MSXML2.ServerXMLHTTP
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.Open "POST", apiUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.send jsonPayload
    responseText = xmlHttp.responseText
#End If
```

---

## Step-by-Step Instructions to Run and Verify

### Step 1: Install Dependencies
Open your terminal and navigate to the assignment directory:
```bash
cd "Assignment for (Productionisation of Machine Learning System)/Final Exam (Assignment)"
pip install -r requirements.txt
```

### Step 2: Start the FastAPI Server
Launch the backend server via Uvicorn:
```bash
uvicorn main:app --reload --port 8000
```
- The interactive Swagger docs will be live at: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- The browser tester will be available at: [http://127.0.0.1:8000/interactive-tester](http://127.0.0.1:8000/interactive-tester)

### Step 3: Open the Excel Workbook
1. Open [`Bank_Loan_Default_Prediction.xlsm`](file:///Users/samson/Desktop/Data%20Experimention%20Folders/Data%20Science%20Institute%20%28Courses%29/All%20Assignment%20for%20Course%20Work/Assignment%20for%20%28Productionisation%20of%20Machine%20Learning%20System%29/Final%20Exam%20%28Assignment%29/Bank_Loan_Default_Prediction.xlsm) in Microsoft Excel.
2. If prompted, click "Enable Macros" or "Enable Content".
3. On the "Single Customer Prediction" sheet, enter or edit customer attributes in cells `C6:C11`.
4. Click the blue "Submit to FastAPI (/predict)" button.
5. The predicted default probability, risk tier, and credit recommendation will update immediately in cells `C14:C18` accompanied by an evaluation popup alert!
6. Switch to the "BankLoan_Test Predictions" sheet and click "Run Batch API Prediction" to refresh predictions for all 10 test records from the live API.
