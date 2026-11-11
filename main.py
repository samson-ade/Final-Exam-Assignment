import os
from typing import List, Optional
import numpy as np

import pandas as pd
import joblib
from pydantic import BaseModel, Field

from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse

# ---------------------------------------------------------------------------
# Application Initialization & Setup
# ---------------------------------------------------------------------------
app = FastAPI(
    title= "Bank Loan Default Prediction API",
    description= (
        "Production-grade REST API powered by a 500-tree Random Forest Classifier "
        "to predict loan default probability for retail banking customers. "
        "Designed for seamless integration with Microsoft Excel (.xlsm) and enterprise frontends."
    ),
    version= "1.0.0",
    docs_url= "/docs",
    redoc_url= "/redoc"
)

# Enable CORS for cross-origin requests from Excel Web queries and browser clients
app.add_middleware(
    CORSMiddleware,
    allow_origins= ["*"],
    allow_credentials= True,
    allow_methods= ["*"],
    allow_headers= ["*"],
)

# ---------------------------------------------------------------------------
# Model Loading & Verification
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_PATH = os.path.join(BASE_DIR, "model.joblib")
TEST_DATA_PATH = os.path.join(BASE_DIR, "BANK LOAN_TEST.csv")
TRAIN_DATA_PATH = os.path.join(BASE_DIR, "BANK LOAN.csv")

FEATURE_NAMES = ["AGE", "EMPLOY", "ADDRESS", "DEBTINC", "CREDDEBT", "OTHDEBT"]

model = None
if os.path.exists(MODEL_PATH):
    model = joblib.load(MODEL_PATH)
else:
    # Train and save model if artifact is missing
    if os.path.exists(TRAIN_DATA_PATH):
        from sklearn.ensemble import RandomForestClassifier
        train_data = pd.read_csv(TRAIN_DATA_PATH)

        X = train_data[FEATURE_NAMES]
        y = train_data["DEFAULTER"]

        model = RandomForestClassifier(n_estimators= 500, random_state= 42)
        model.fit(X, y)
        joblib.dump(model, MODEL_PATH)


# ---------------------------------------------------------------------------
# Data Schemas (Pydantic Models)
# ---------------------------------------------------------------------------
class CustomerLoanFeatures(BaseModel):
    AGE: float = Field(
        ...,
        description= "Age category or age of customer (1= <30, 2= 30-45, 3= >45)",
        examples= [2.0]
    )
    EMPLOY: float = Field(
        ...,
        description= "Years with current employer",
        ge= 0.0,
        examples= [8.0]
    )
    ADDRESS: float = Field(
        ...,
        description= "Years at current address / residential tenure",
        ge= 0.0,
        examples= [5.0]
    )
    DEBTINC: float = Field(
        ...,
        description= "Debt-to-Income ratio (percentage, e.g. 17.0 for 17%)",
        ge= 0.0,
        examples= [17.0]
    )
    CREDDEBT: float = Field(
        ...,
        description= "Credit card debt in thousands ($k)",
        ge= 0.0,
        examples= [1.46]
    )
    OTHDEBT: float = Field(
        ...,
        description= "Other debt in thousands ($k)",
        ge= 0.0,
        examples= [4.00]
    )


class LoanPredictionResponse(BaseModel):
    default_probability: float = Field(..., description= "Estimated probability of default (0.0 to 1.0)")
    default_probability_pct: str = Field(..., description= "Formatted probability percentage")
    prediction: int = Field(..., description= "Binary classification (1 = Default, 0 = Non-Default)")
    prediction_label: str = Field(..., description= "Human-readable prediction label")
    risk_tier: str = Field(..., description= "Risk tier: Low, Moderate, or High")
    recommendation: str = Field(..., description= "Credit decision recommendation")
    input_features: dict = Field(..., description= "Echo of input features evaluated")


class BatchCustomerLoanFeatures(BaseModel):
    customers: List[CustomerLoanFeatures]


# ---------------------------------------------------------------------------
# Helper Decision Function
# ---------------------------------------------------------------------------
def assess_credit_risk(prob: float) -> tuple[str, str, str, int]:
    """Classifies probability into operational credit risk tiers and decisions."""
    if prob >= 0.50:
        return "Defaulter", "High Risk", "Decline / Reject Application", 1
    elif prob >= 0.30:
        return "Non-Defaulter", "Moderate Risk", "Manual Underwriter Review Required", 0
    else:
        return "Non-Defaulter", "Low Risk", "Approve Loan Application", 0


# ---------------------------------------------------------------------------
# API Endpoints
# ---------------------------------------------------------------------------
@app.get("/", summary= "API Root & Overview")
def home():
    return {
        "title": "Bank Loan Default Prediction API",
        "version": "1.0.0",
        "status": "online",
        "model": "Random Forest Classifier (500 trees)",
        "features_required": FEATURE_NAMES,
        "endpoints": {
            "/predict": "POST - Predict loan default probability for a single customer (Excel VBA compatible)",
            "/predict-batch": "POST - Batch inference for multiple customer profiles",
            "/test-predictions": "GET - Pre-computed test set predictions from BankLoan_Test",
            "/health": "GET - API health check and model status",
            "/interactive-tester": "GET - Browser GUI for testing customer applications",
            "/docs": "GET - Interactive Swagger documentation"
        },
        "excel_vba_support": "Compatible with WinHTTP, MSXML2.ServerXMLHTTP (Windows), and curl MacScript (macOS)."
    }


@app.get("/health", summary= "Health Check")
def health_check():
    if model is None:
        raise HTTPException(
            status_code= status.HTTP_503_SERVICE_UNAVAILABLE,
            detail= "Model is not loaded."
        )
    return {
        "status": "healthy",
        "model_loaded": True,
        "model_type": str(type(model).__name__),
        "n_estimators": getattr(model, "n_estimators", 500),
        "features_expected": FEATURE_NAMES
    }


@app.post("/predict", response_model= LoanPredictionResponse, summary= "Predict Loan Default Risk")
def predict_default(customer: CustomerLoanFeatures):
    """
    Primary endpoint for loan default probability estimation.
    Accepts customer profile parameters and returns estimated default probability,
    risk category, and recommended credit action.
    """
    if model is None:
        raise HTTPException(
            status_code= status.HTTP_503_SERVICE_UNAVAILABLE,
            detail= "Random Forest model artifact is not loaded."
        )

    # Format features as DataFrame with correct column names
    input_dict = {
        "AGE": [customer.AGE],
        "EMPLOY": [customer.EMPLOY],
        "ADDRESS": [customer.ADDRESS],
        "DEBTINC": [customer.DEBTINC],
        "CREDDEBT": [customer.CREDDEBT],
        "OTHDEBT": [customer.OTHDEBT]
    }
    input_df = pd.DataFrame(input_dict)

    # Generate probabilities
    try:
        proba = float(model.predict_proba(input_df)[0, 1])
    except Exception as e:
        raise HTTPException(
            status_code= status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail= f"Inference error: {str(e)}"
        )

    prob_rounded = round(proba, 4)
    pred_label, risk_tier, recommendation, binary_pred = assess_credit_risk(prob_rounded)

    return LoanPredictionResponse(
        default_probability= prob_rounded,
        default_probability_pct= f"{prob_rounded * 100:.2f}%",
        prediction= binary_pred,
        prediction_label= pred_label,
        risk_tier= risk_tier,
        recommendation= recommendation,
        input_features= customer.model_dump()
    )


@app.post("/predict-batch", summary= "Batch Loan Default Prediction")
def predict_batch(batch: BatchCustomerLoanFeatures):
    """Generates predictions and default probabilities for multiple customer applications."""
    if model is None:
        raise HTTPException(status_code= 503, detail= "Model is not loaded.")

    if not batch.customers:
        return {"predictions": [], "total_records": 0}

    input_records = [c.model_dump() for c in batch.customers]
    input_df = pd.DataFrame(input_records)[FEATURE_NAMES]

    probabilities = model.predict_proba(input_df)[:, 1]

    results = []
    for i, prob in enumerate(probabilities):
        prob_r = round(float(prob), 4)
        pred_label, risk_tier, recommendation, binary_pred = assess_credit_risk(prob_r)

        results.append({
            "index": i + 1,
            "default_probability": prob_r,
            "default_probability_pct": f"{prob_r * 100:.2f}%",
            "prediction": binary_pred,
            "prediction_label": pred_label,
            "risk_tier": risk_tier,
            "recommendation": recommendation,
            "input": input_records[i]
        })

    return {
        "total_records": len(results),
        "predictions": results
    }


@app.get("/test-predictions", summary= "Get BankLoan_Test Predictions")
def get_test_predictions():
    """Returns predictions on the 10 customers from BANK LOAN_TEST.csv."""
    if not os.path.exists(TEST_DATA_PATH):
        raise HTTPException(status_code= 404, detail= "BANK LOAN_TEST.csv not found.")

    test_df = pd.read_csv(TEST_DATA_PATH)
    X_test = test_df[FEATURE_NAMES]
    probs = model.predict_proba(X_test)[:, 1]

    test_records = []
    for idx, row in test_df.iterrows():
        p = round(float(probs[idx]), 4)
        pred_label, risk_tier, rec, binary_pred = assess_credit_risk(p)
        actual = int(row["DEFAULTER"]) if "DEFAULTER" in row else None
        test_records.append({
            "SN": int(row["SN"]),
            "AGE": float(row["AGE"]),
            "EMPLOY": float(row["EMPLOY"]),
            "ADDRESS": float(row["ADDRESS"]),
            "DEBTINC": float(row["DEBTINC"]),
            "CREDDEBT": float(row["CREDDEBT"]),
            "OTHDEBT": float(row["OTHDEBT"]),
            "ACTUAL_DEFAULTER": actual,
            "PREDICTED_PROBABILITY": p,
            "PREDICTED_PROBABILITY_PCT": f"{p * 100:.2f}%",
            "PREDICTED_CLASS": binary_pred,
            "RISK_TIER": risk_tier,
            "RECOMMENDATION": rec,
            "CORRECT": (binary_pred == actual) if actual is not None else None
        })

    return {
        "dataset": "BANK LOAN_TEST.csv",
        "sample_count": len(test_records),
        "accuracy": sum(1 for r in test_records if r["CORRECT"]) / len(test_records) if test_records else 0,
        "results": test_records
    }


@app.get("/interactive-tester", response_class= HTMLResponse, summary= "Interactive Web GUI")
def interactive_tester():
    """Provides a sleek browser-based testing interface for the API."""
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bank Loan Default Risk Tester</title>
    <style>
        :root {
            --primary: #1e3a8a;
            --primary-hover: #1e40af;
            --bg: #f8fafc;
            --card: #ffffff;
            --border: #e2e8f0;
            --text: #0f172a;
            --text-muted: #64748b;
            --success: #16a34a;
            --warning: #d97706;
            --danger: #dc2626;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
            background-color: var(--bg);
            color: var(--text);
            padding: 30px 15px;
            display: flex;
            justify-content: center;
        }
        .container {
            max-width: 800px;
            width: 100%;
        }
        .header {
            text-align: center;
            margin-bottom: 25px;
        }
        .header h1 {
            color: var(--primary);
            font-size: 26px;
            margin-bottom: 8px;
        }
        .header p {
            color: var(--text-muted);
            font-size: 14px;
        }
        .card {
            background: var(--card);
            border-radius: 12px;
            padding: 24px;
            border: 1px solid var(--border);
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
            margin-bottom: 20px;
        }
        .form-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }
        .form-group label {
            display: block;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 6px;
            color: #334155;
        }
        .form-group input, .form-group select {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            transition: border-color 0.2s;
        }
        .form-group input:focus, .form-group select:focus {
            border-color: var(--primary);
        }
        .btn-submit {
            width: 100%;
            background-color: var(--primary);
            color: white;
            padding: 12px;
            border: none;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            transition: background-color 0.2s;
        }
        .btn-submit:hover {
            background-color: var(--primary-hover);
        }
        .result-box {
            display: none;
            margin-top: 20px;
            padding: 20px;
            border-radius: 8px;
            background: #f1f5f9;
            border-left: 6px solid var(--primary);
        }
        .metric-row {
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #e2e8f0;
            font-size: 14px;
        }
        .metric-row:last-child { border-bottom: none; }
        .prob-large {
            font-size: 32px;
            font-weight: 700;
            text-align: center;
            margin: 10px 0;
        }
        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 9999px;
            font-weight: 600;
            font-size: 12px;
        }
        .badge-danger { background: #fee2e2; color: var(--danger); }
        .badge-warning { background: #fef3c7; color: var(--warning); }
        .badge-success { background: #dcfce7; color: var(--success); }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Bank Loan Default Risk Assessor</h1>
            <p>FastAPI Interactive Client &bull; 500-Tree Random Forest Classifier</p>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px; font-size: 16px;">Customer Financial & Demographic Profile</h3>
            <form id="loanForm">
                <div class="form-grid">
                    <div class="form-group">
                        <label for="AGE">Age Category (AGE)</label>
                        <select id="AGE" required>
                            <option value="1">1: Under 30</option>
                            <option value="2" selected>2: 30 to 45</option>
                            <option value="3">3: Over 45</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="EMPLOY">Employment Tenure (Years)</label>
                        <input type="number" id="EMPLOY" value="8" step="1" min="0" required>
                    </div>
                    <div class="form-group">
                        <label for="ADDRESS">Address Tenure (Years)</label>
                        <input type="number" id="ADDRESS" value="5" step="1" min="0" required>
                    </div>
                    <div class="form-group">
                        <label for="DEBTINC">Debt-to-Income Ratio (%)</label>
                        <input type="number" id="DEBTINC" value="17.0" step="0.1" min="0" required>
                    </div>
                    <div class="form-group">
                        <label for="CREDDEBT">Credit Card Debt ($ thousands)</label>
                        <input type="number" id="CREDDEBT" value="1.46" step="0.01" min="0" required>
                    </div>
                    <div class="form-group">
                        <label for="OTHDEBT">Other Debt ($ thousands)</label>
                        <input type="number" id="OTHDEBT" value="4.00" step="0.01" min="0" required>
                    </div>
                </div>
                <button type="submit" class="btn-submit">Submit to FastAPI (/predict)</button>
            </form>

            <div id="resultBox" class="result-box">
                <div style="text-align: center; font-size: 13px; color: var(--text-muted); text-transform: uppercase;">Estimated Default Probability</div>
                <div id="probDisplay" class="prob-large">--%</div>
                <div style="text-align: center; margin-bottom: 15px;">
                    <span id="tierBadge" class="badge">--</span>
                </div>
                <div class="metric-row">
                    <span>Classification:</span>
                    <strong id="labelDisplay">--</strong>
                </div>
                <div class="metric-row">
                    <span>Decision Recommendation:</span>
                    <strong id="recDisplay">--</strong>
                </div>
                <div class="metric-row">
                    <span>Endpoint Called:</span>
                    <code>POST /predict</code>
                </div>
            </div>
        </div>
    </div>

    <script>
        document.getElementById('loanForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            const payload = {
                AGE: parseFloat(document.getElementById('AGE').value),
                EMPLOY: parseFloat(document.getElementById('EMPLOY').value),
                ADDRESS: parseFloat(document.getElementById('ADDRESS').value),
                DEBTINC: parseFloat(document.getElementById('DEBTINC').value),
                CREDDEBT: parseFloat(document.getElementById('CREDDEBT').value),
                OTHDEBT: parseFloat(document.getElementById('OTHDEBT').value)
            };

            try {
                const response = await fetch('/predict', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await response.json();

                const box = document.getElementById('resultBox');
                box.style.display = 'block';

                document.getElementById('probDisplay').innerText = data.default_probability_pct;
                document.getElementById('labelDisplay').innerText = data.prediction_label;
                document.getElementById('recDisplay').innerText = data.recommendation;

                const badge = document.getElementById('tierBadge');
                badge.innerText = data.risk_tier;
                if (data.default_probability >= 0.50) {
                    badge.className = 'badge badge-danger';
                    document.getElementById('probDisplay').style.color = 'var(--danger)';
                    box.style.borderLeftColor = 'var(--danger)';
                } else if (data.default_probability >= 0.30) {
                    badge.className = 'badge badge-warning';
                    document.getElementById('probDisplay').style.color = 'var(--warning)';
                    box.style.borderLeftColor = 'var(--warning)';
                } else {
                    badge.className = 'badge badge-success';
                    document.getElementById('probDisplay').style.color = 'var(--success)';
                    box.style.borderLeftColor = 'var(--success)';
                }
            } catch (err) {
                alert('API Request Failed: ' + err.message);
            }
        });
    </script>
</body>
</html>
"""
    return html_content


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host= "127.0.0.1", port= 8000, reload= True)
