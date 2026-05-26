# Network Intrusion Detection with Machine Learning — Big Data Security

> **Graduate Course Project** | Big Data | UNSW-NB15 Dataset

---

## Project Overview

This project builds a **network intrusion detection system** using machine learning on the UNSW-NB15 dataset. Network traffic data exhibits core Big Data characteristics — high **volume**, **velocity**, and **variety** — making manual analysis infeasible. A Random Forest classifier is trained to distinguish normal traffic from attacks, and results are evaluated both technically and from an **ethical/privacy perspective**.

The project goes beyond standard model evaluation by analyzing:
- The **security risk** of missed attacks (False Negatives)
- The **ethical risk** of flagging innocent users as threats (False Positives)
- The **privacy implications** of processing network behavior data under GDPR principles

---

## Dataset

**UNSW-NB15** — Created by the Australian Centre for Cyber Security (ACCS) using the IXIA PerfectStorm tool.

| Property | Value |
|---|---|
| Total records | 2,540,044 |
| Features | 49 |
| Training set | 82,332 records |
| Test set | 175,341 records |
| Attack categories | 9 |

**Attack types:** Fuzzers, DoS, Backdoor, Exploits, Reconnaissance, Shellcode, Worms, Analysis, Generic

> Dataset source: [Kaggle — UNSW-NB15](https://www.kaggle.com/datasets/mrwellsdavid/unsw-nb15)  
> The CSV files are **not included** in this repository due to size. Download from the link above and place in `data/`.

---

## Results

| Metric | Value |
|---|---|
| **Accuracy** | **90.13%** |
| **F1-Score** | **0.9225** |
| **ROC-AUC** | **0.9864** |
| Precision (Attack) | 99.0% |
| Recall (Attack) | 86.2% |

### Confusion Matrix Breakdown

| Category | Count | Interpretation |
|---|---|---|
| True Positive (TP) | 103,059 | Attacks correctly detected |
| True Negative (TN) | 54,977 | Normal traffic correctly classified |
| **False Negative (FN)** | **16,282** | ⚠ Real attacks **missed** → Security Risk |
| **False Positive (FP)** | **1,023** | ⚠ Innocent users flagged → **Ethical Risk** |

---

## Ethical & Privacy Analysis

A key contribution of this project is framing model errors as **ethical and operational risks**, not just accuracy numbers:

- **False Positives** → Innocent users are incorrectly labeled as attackers, leading to unjust access denial and trust erosion (ethical/fairness risk)
- **False Negatives** → Real intrusions go undetected, exposing the network to breaches (security risk)
- **Privacy (GDPR)** → Network traffic features such as IP addresses and connection timestamps may qualify as personal data. Processing them requires compliance with principles of **purpose limitation**, **data minimization**, and **proportionality**

---

## Methodology

```
Raw Data (CSV)
     │
     ▼
Preprocessing
  • Label Encoding (proto, service, state)
  • StandardScaler normalization
  • Median imputation for missing values
     │
     ▼
Model Training
  • Random Forest: 100 trees, max_depth=20
  • Class weighting (balanced) for imbalance
     │
     ▼
Evaluation
  • Accuracy, F1, ROC-AUC
  • Confusion Matrix (raw + normalized)
  • Feature Importance
     │
     ▼
Ethical Analysis
  • FP/FN risk interpretation
  • GDPR & privacy review
```

---

## Top 5 Most Important Features

| Feature | Description |
|---|---|
| `ct_state_ttl` | Connection state / TTL combination count |
| `sttl` | Source-to-destination TTL value |
| `dttl` | Destination-to-source TTL value |
| `ct_dst_src_ltm` | Recent connection count (dst→src) |
| `rate` | Packet rate of the connection |

---

## Project Files

| File | Description |
|---|---|
| `unsw_nb15_intrusion_detection.ipynb` | Full Jupyter notebook — data loading, preprocessing, modeling, evaluation, ethics |
| `run_analysis.py` | Standalone script to reproduce all charts |
| `create_presentation.py` | Script that generates the PowerPoint presentation |
| `BigVeri_SaldiriTespiti_Sunum.pptx` | 14-slide academic presentation (Turkish) with speaker notes |
| `*.png` | All generated charts (confusion matrix, ROC, feature importance, etc.) |
| `data/` | *(Not included)* — place UNSW-NB15 CSV files here |

---

## Setup & Run

### 1. Create the environment

```bash
conda create -n bigdataproje python=3.11
conda activate bigdataproje
pip install pandas numpy matplotlib seaborn scikit-learn notebook ipykernel
```

### 2. Add the dataset

Download from Kaggle and place in the `data/` folder:
```
data/
  UNSW_NB15_training-set.csv
  UNSW_NB15_testing-set.csv
```

### 3. Run the notebook

```bash
jupyter notebook unsw_nb15_intrusion_detection.ipynb
```

Or regenerate all charts:

```bash
python run_analysis.py
```

---

## Key Takeaway

> *"Big data security systems should be designed not only for technical accuracy, but also in alignment with ethical, fairness, and privacy principles."*

The precision–recall trade-off is not merely a modeling decision — it carries real consequences for both security and human rights. Tuning the decision threshold should be informed by the deployment context and the acceptable risk balance.

---

## References

- Moustafa, N. & Slay, J. (2015). *UNSW-NB15: A comprehensive data set for network intrusion detection systems.* MilCIS.
- European Commission — GDPR: Regulation (EU) 2016/679
- Scikit-learn documentation: [scikit-learn.org](https://scikit-learn.org)
