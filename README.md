# 💳 신용카드 이상거래 탐지 (Credit Card Fraud Detection)

> 머신러닝 기반 불균형 이진 분류 프로젝트  
> 전체 거래의 **0.17%** 만 사기인 극심한 불균형 데이터에서 사기 거래를 탐지합니다.

---

## 📌 프로젝트 개요

신용카드 거래 데이터에서 사기(Fraud) 거래를 정확히 탐지하는 분류 모델을 구축합니다.

**핵심 문제:**  
단순히 "전부 정상"이라고 예측해도 정확도 99.8%가 나오는 불균형 데이터  
→ Accuracy 지표는 의미가 없으며, **F1-Score · AUC-ROC · Recall** 기준으로 모델을 평가해야 합니다.

---

## 📂 프로젝트 구조

```
fraud_detection/
├── notebooks/
│   └── fraud_detection.ipynb          # 메인 분석 노트북 (전체 파이프라인)
├── data/
│   └── creditcard.csv                 # Kaggle에서 다운로드 (아래 안내 참고)
├── 신용카드_이상거래_탐지.pptx           # 프로젝트 발표 PPT
├── requirements.txt
└── README.md
```

---

## 🗃️ 데이터

| 항목 | 내용 |
|------|------|
| 출처 | [Kaggle - Credit Card Fraud Detection](https://www.kaggle.com/datasets/mlg-ulb/creditcardfraud) |
| 크기 | 284,807건 거래 |
| 사기 건수 | 492건 (0.17%) |
| 특성 | V1~V28 (PCA 익명화), Amount, Time |
| 타겟 | Class (0=정상, 1=사기) |

> `data/` 폴더에 `creditcard.csv`를 직접 다운로드해서 넣어주세요.

---

## ⚙️ 전처리 파이프라인

1. **정규화** — Amount · Time을 StandardScaler로 변환 (V1~V28은 이미 PCA 정규화됨)
2. **훈련/테스트 분리** — 80:20 비율, stratify 옵션으로 사기 비율 유지
3. **SMOTE** — 훈련 데이터 내 사기 394건 → 227,451건으로 합성 증강  
   *(테스트 데이터에는 적용 안 함)*

---

## 🤖 사용 모델 (총 9개)

### 기초 모델
| 모델 | 특징 |
|------|------|
| 로지스틱 회귀 | 확률 기반 판정, 빠르고 해석 쉬움 |
| 결정 트리 | 조건 질문 트리 구조, 직관적 |
| KNN | 유사 거래 K개 다수결 판정 |
| 나이브 베이즈 | 확률 통계 기반, 매우 빠름 |
| SVM | 최적 경계선 탐색, 고차원에 강함 |

### 앙상블 모델
| 모델 | 특징 |
|------|------|
| 랜덤 포레스트 | 결정 트리 100개 다수결, 안정적 |
| XGBoost | 오류 집중 반복 학습, 캐글 최강 알고리즘 |
| LightGBM | XGBoost 대비 빠른 속도, 금융권 실무 표준 |

### 딥러닝 모델
| 모델 | 특징 |
|------|------|
| 신경망 (MLP) | 64→32 은닉층, early_stopping 적용 |

---

## 📊 성능 결과

| 순위 | 모델 | F1-Score | AUC-ROC | Recall |
|------|------|---------|---------|--------|
| 🥇 | KNN | **0.859** | 0.944 | 0.806 |
| 🥈 | 랜덤 포레스트 | 0.821 | 0.969 | 0.816 |
| 🥉 | 신경망(MLP) | 0.760 | 0.970 | 0.806 |
| 4 | LightGBM | 0.640 | 0.969 | 0.888 |
| 5 | XGBoost | 0.602 | **0.978** | 0.857 |
| 6 | SVM | 0.218 | 0.977 | 0.898 |
| 7 | 결정 트리 | 0.144 | 0.895 | 0.806 |
| 8 | 나이브 베이즈 | 0.110 | 0.963 | 0.847 |
| 9 | 로지스틱 회귀 | 0.109 | 0.970 | 0.918 |

> **SVM** 은 1만 건 샘플 학습으로 인한 성능 제약 있음 (전체 학습 시 성능 향상 예상)  
> **XGBoost/LightGBM** 은 하이퍼파라미터 튜닝 시 F1 0.85+ 달성 가능

---

## 💡 핵심 인사이트

- **불균형 데이터에서는 Accuracy가 의미 없다.** F1-Score와 Recall로 평가해야 한다.
- **앙상블 모델(RF, XGB, LGBM)이 기초 모델 대비 전반적으로 우수**하지만, 하이퍼파라미터 튜닝 없이는 기초 모델(KNN)에 역전될 수 있다.
- **SMOTE는 강력하지만 주의가 필요하다.** 테스트 데이터에 적용하면 현실과 다른 왜곡된 평가가 나온다.
- **사기 탐지에서는 Recall(놓치지 않기)이 Precision(오탐 줄이기)보다 우선순위가 높다.**

---

## 🚀 실행 방법

```bash
# 1. 패키지 설치
pip install -r requirements.txt

# 2. Kaggle에서 데이터 다운로드 후 data/creditcard.csv 로 저장

# 3. 노트북 실행
jupyter notebook notebooks/fraud_detection.ipynb
```

---

## 🛠️ 기술 스택

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Scikit-learn](https://img.shields.io/badge/Scikit--learn-1.4+-orange)
![XGBoost](https://img.shields.io/badge/XGBoost-latest-red)
![LightGBM](https://img.shields.io/badge/LightGBM-latest-green)
![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-orange)

- **분석**: Pandas, NumPy
- **모델**: Scikit-learn, XGBoost, LightGBM
- **불균형 처리**: Imbalanced-learn (SMOTE)
- **시각화**: Matplotlib, Seaborn
