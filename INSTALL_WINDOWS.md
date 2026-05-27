# MCAL — Windows 설치 매뉴얼

**Multi-Channel Sensor Calibration Tool**
대상 센서: PT100 / PT1000 / Strain Gauge 350Ω (16채널)

---

## 목차

1. [시스템 요구사항](#1-시스템-요구사항)
2. [방법 A — 실행 파일(.exe) 다운로드 (권장)](#2-방법-a--실행-파일exe-다운로드-권장)
3. [방법 B — Python 소스 실행](#3-방법-b--python-소스-실행)
4. [첫 실행 방법](#4-첫-실행-방법)
5. [폴더 구조](#5-폴더-구조)
6. [문제 해결](#6-문제-해결)

---

## 1. 시스템 요구사항

| 항목 | 최소 사양 |
|------|----------|
| 운영체제 | Windows 10 / 11 (64-bit) |
| RAM | 4 GB 이상 |
| 디스크 | 500 MB 이상 (실행 파일 포함) |
| 화면 해상도 | 1280 × 720 이상 권장 |

---

## 2. 방법 A — 실행 파일(.exe) 다운로드 (권장)

별도 Python 설치 없이 바로 사용할 수 있습니다.

### 2-1. 다운로드

1. GitHub 저장소 접속: [https://github.com/lhyunso/Calibration](https://github.com/lhyunso/Calibration)
2. 우측 **Releases** 섹션 클릭
3. 최신 릴리즈의 **`MCAL.exe`** 클릭하여 다운로드

> 릴리즈가 없는 경우 **Actions** 탭 → 최신 워크플로우 실행 → **Artifacts** 섹션에서 `MCAL_windows` 다운로드

### 2-2. 설치

별도 설치 과정 없습니다. `MCAL.exe` 파일 하나만 원하는 위치에 저장 후 실행합니다.

> **첫 실행 시 5~10초 소요됩니다** — 내장된 런타임을 임시 폴더에 압축 해제하는 과정이며, 이후 실행은 빨라집니다.

> **Windows 보안 경고가 뜨는 경우:**
> `추가 정보` → `실행` 클릭 (개발자 서명이 없는 경우 표시될 수 있습니다)

---

## 3. 방법 B — Python 소스 실행

소스 코드를 직접 실행하거나 수정이 필요한 경우 사용합니다.

### 3-1. Python 3.11 설치

1. [https://www.python.org/downloads/](https://www.python.org/downloads/) 접속
2. **Python 3.11.x** (Windows installer 64-bit) 다운로드
3. 설치 시 **"Add Python to PATH"** 반드시 체크

![Python 설치 시 PATH 체크 필수](https://www.python.org/static/community_logos/python-logo.png)

4. 설치 확인:
```cmd
python --version
```
출력 예시: `Python 3.11.9`

### 3-2. 소스 코드 다운로드

**Git 사용 시:**
```cmd
git clone https://github.com/lhyunso/Calibration.git
cd Calibration
```

**Git 미사용 시:**
1. GitHub 저장소 → **Code** → **Download ZIP**
2. 압축 해제 후 폴더 진입

### 3-3. 패키지 설치

```cmd
pip install -r requirements.txt
```

설치되는 주요 패키지:

| 패키지 | 용도 |
|--------|------|
| customtkinter | GUI 프레임워크 |
| openpyxl | xlsx 출력 |
| python-docx | docx 출력 |
| reportlab | pdf 출력 |
| matplotlib / numpy | 차트 생성 |

### 3-4. 실행

```cmd
python src\gui.py
```

---

## 4. 첫 실행 방법

### 4-1. 센서 타입 선택

실행 후 상단에서 센서 타입을 선택합니다.

| 센서 타입 | R_nom | 기본 Inst.Amp.Gain |
|-----------|-------|-------------------|
| PT100 | 100 Ω | 10 |
| PT1000 | 1000 Ω | 10 |
| Strain 350Ω | 350 Ω | 100 |

### 4-2. CSV 파일 로드

1. **CSV 폴더 선택** 버튼 클릭
2. 모사저항별 CSV 파일이 있는 폴더 선택

CSV 파일명 형식 (저항값이 파일명에 포함되어야 합니다):
```
ANA.BRI.Q100_Calibration data_80.csv
ANA.BRI.Q100_Calibration data_90.csv
ANA.BRI.Q100_Calibration data_100.csv
ANA.BRI.Q100_Calibration data_110.csv
ANA.BRI.Q100_Calibration data_120.csv
```

### 4-3. 메타데이터 입력

| 항목 | 설명 |
|------|------|
| 모듈명 | 교정 대상 모듈 이름 |
| 모델명 | 제품 모델명 |
| S/N | 시리얼 번호 |
| 교정일자 | YYYY-MM-DD 형식 |
| Inst.Amp.Gain | 계측 증폭기 이득 (센서 선택 시 자동 입력) |

### 4-4. 캘리브레이션 실행 및 출력

1. **캘리브레이션 실행** 버튼 클릭
2. 출력 형식 선택: xlsx (기본) / docx / pdf
3. **저장** 버튼 클릭 → 출력 파일 생성

**출력 파일 구성 (xlsx 기준):**

| 시트 | 내용 |
|------|------|
| Cover_Summary | 문서 정보 + 전 채널 요약 (Gain, Offset, 합불) |
| CH01 ~ CH16 | 채널별 Decimal/전압/저항/편차 + 차트 2개 |
| 3W_Reference | 모사저항 기준 전압 레퍼런스 표 |
| Raw_Data | 원본 ADC Decimal 값 |

---

## 5. 폴더 구조

`MCAL.exe` 실행 시 같은 위치에 아래 항목이 자동 생성됩니다.

```
(원하는 위치)\
  MCAL.exe           ← 단일 실행 파일 (이것만 배포)
  outputs\           ← 출력 파일 저장 위치 (자동 생성)
    xlsx\
    docx\
    pdf\
  settings.json      ← 마지막 설정 저장 (자동 생성)
```

---

## 6. 문제 해결

### 실행 시 "vcruntime140.dll 없음" 오류

Microsoft Visual C++ 재배포 패키지 설치 필요:
- [vc_redist.x64.exe 다운로드](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### 실행 시 화면이 흰색으로 나오는 경우

Python 버전 문제입니다. **Python 3.11** 이상을 사용해야 합니다 (방법 B 해당).

```cmd
python --version   ← 3.11.x 인지 확인
```

### CSV 파일을 인식하지 못하는 경우

파일명에 저항값(숫자)이 포함되어 있는지 확인합니다.
- ✅ `calibration_100.csv`
- ✅ `data_100ohm.csv`
- ❌ `calibration_nominal.csv`

### 출력 파일이 열리지 않는 경우

이미 동일한 파일이 Excel/Word에서 열려 있는 경우 닫은 후 재시도합니다.

### 기타 오류

GitHub Issues에 오류 메시지와 함께 문의해 주세요:
[https://github.com/lhyunso/Calibration/issues](https://github.com/lhyunso/Calibration/issues)

---

*MCAL v1.0 — ANA.BRI.Q100 Multi-Channel Sensor Calibration Tool*
