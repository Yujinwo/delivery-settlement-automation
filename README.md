## 🧩 데이터 정산 자동화 프로젝트

CSV 주문 데이터를 기반으로 가맹점 정산을 자동 계산하고  
엑셀 보고서를 생성하는 Python 자동화 프로젝트입니다.

---

## 🧩개요

반복적으로 수행되는 정산 업무를 자동화하여  
데이터 취합 → 정제 → 정산 계산 → 엑셀 리포트 생성까지  
원클릭으로 처리하도록 구현했습니다.

---

## ⚙️ 주요 기능

- 여러 CSV 주문 파일 자동 취합
- 데이터 전처리
  - 결측치 제거
  - 이상치 검증
  - 중복 주문 제거
- 정산 금액 자동 계산
- 가맹점별 / 날짜별 매출 집계
- openpyxl 기반 엑셀 서식 자동 적용
- 입력/출력 폴더 기반 자동화 처리

---

## 🛠 사용 기술

- Python
- pandas
- openpyxl
- glob / os

---

### 1️⃣ 패키지 설치

```bash
pip install pandas openpyxl
```
### 2️⃣ input 폴더에 CSV 파일 배치

 필수 컬럼:
 - order_id
 - order_date
 - store_name
 - menu_name
 - qty
 - unit_price
 - fee_rate

3️⃣ 실행

```bash
python main.py
```

## 📷 실행 결과

### 1. 원본 주문 데이터
![raw](https://drive.google.com/uc?id=1borjMEecW6dWpF0ErpJxMp8fd0AbPo7M)

### 2. 자동화 실행
![run](https://drive.google.com/uc?id=1UA0ydrmYeC46tzAED2w0xeMaGFoyb9bV)

### 3. 가맹점별 정산 결과
![store](https://drive.google.com/uc?id=1aJsdx2pbJfo2Z2OP8FKZ5f2o-Ym4X_xE)

### 4. 날짜별 매출 집계
![date](https://drive.google.com/uc?id=1VyoV-kwCGnO3kjWE4zm33JaaXw6T823o)