# SpDocumentsConverter

성풍 출고장과 토글 주문을 이카운트·위하고 등의 Excel 업로드 양식으로 변환하는 데스크톱 프로그램입니다.
AMD64 Windows와 Apple Silicon Mac을 대상으로 같은 소스와 PyInstaller 설정을 사용합니다.
이 브랜치는 기존 화면, 기능, 입출력 및 변환 규칙을 유지하면서 운영체제 호환성을 정리합니다.

## 실행 환경

- Python 3.14와 Tcl/Tk. `tkinter`는 표준 라이브러리지만 Python 설치에 Tcl/Tk가 포함되어 있어야 합니다.
- 저장된 Excel 파일을 읽는 기능은 `openpyxl`을 사용합니다.
- **현재 시트 / 선택 영역** 기능은 해당 컴퓨터에 설치된 Microsoft Excel을 `xlwings`로 제어합니다.
  `pip`가 Windows에서는 `pywin32`, macOS에서는 `appscript`와 `psutil`을 자동 설치합니다.
  Excel 추가 기능(add-in)은 필요하지 않습니다.

Mac의 Excel 연동은 처음 사용할 때 자동화 권한을 허용해야 합니다.
거부했다면 시스템 설정 → 개인정보 보호 및 보안 → 자동화에서 이 앱(소스 실행 시 터미널/Python)의 Excel 제어를 허용합니다.

## 개발 환경 준비

Windows에서는 AMD64(64비트) Python을 설치하고 Tcl/Tk를 포함합니다. PowerShell에서:

```powershell
py -3.14 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-build.txt
.venv\Scripts\python main.py
```

Apple Silicon Mac에서 Homebrew Python을 사용하는 경우:

```sh
brew install python@3.14 python-tk@3.14
python3.14 -m venv .venv
.venv/bin/python -m pip install -r requirements-build.txt
.venv/bin/python main.py
```

`python -m tkinter`로 Tcl/Tk 설치를 확인할 수 있습니다.
아래 명령의 `python`은 위에서 준비한 `.venv`의 Python을 사용합니다.

## 사용

- **첫 번째 출고장 탭**: Excel에서 문서를 열고 현재 시트 또는 선택 영역을 변환합니다.
  선택 영역 방식에는 제목 한 행을 가리키는 `TITLES` 이름 범위가 필요합니다.
- **두 번째 출고장 탭**: 파일과 시트를 선택한 뒤 실행합니다.
- **토글**: Excel의 현재 시트로 실행합니다. 기존 파일/시트 선택 UI의 동작도 유지합니다.

파일 선택 방식은 Excel이 마지막 저장한 수식 계산값을 읽으므로 원본의 변경 내용을 먼저 저장합니다.
기존 파일 선택창에는 `.xls`도 표시되지만 `openpyxl`로 직접 읽을 수는 없습니다.
구형 `.xls` 파일은 Excel에서 열어 현재 시트 방식으로 변환하거나 `.xlsx`로 저장합니다.
변환 결과는 운영체제의 임시 폴더에 생성되고 기본 연결 프로그램으로 열립니다. 보관할 파일은 원하는 위치에 저장합니다.

## 테스트와 빌드

```sh
python -m unittest discover -s tests -v
python main.py --smoke-test
python -m PyInstaller --noconfirm main.spec
```

- AMD64 Windows 결과: `dist/SpDocumentsConverter/SpDocumentsConverter.exe`.
  배포할 때는 같은 폴더의 `_internal` 등도 함께 전달합니다.
- Mac 결과: `dist/SpDocumentsConverter.app` (arm64).
  Tcl/Tk 및 품목 조회표 두 개가 포함되며 Excel 자동화 권한 설명도 설정됩니다.
- 빌드는 각 운영체제에서 실행합니다. Windows에서 Mac 앱을 빌드하거나 반대로 빌드하지 않습니다.
- Python과 라이브러리는 결과물에 포함되므로 배포받는 사람은 별도로 설치하지 않아도 됩니다.
  현재 시트 / 선택 영역 기능을 쓰는 사람에게는 Microsoft Excel이 필요합니다.

GitHub Actions는 Windows x64와 Mac arm64에서 회귀 테스트, GUI 시작 검사, 빌드 및 패키징 후 시작 검사를 수행하도록 구성되어 있습니다.
Mac 배포 압축에는 `.app`의 실행 권한과 심볼릭 링크를 보존합니다. Developer ID 서명·공증은 이 설정에 포함하지 않습니다.

자동 테스트는 Excel을 제어하지 않습니다. 실제 Excel에서 현재 시트 복사, 선택 영역 읽기,
생성된 파일 열기와 Mac의 최초 권한 요청은 각 운영체제에서 별도로 확인합니다.

## 참고

- [Python tkinter](https://docs.python.org/3/library/tkinter.html)
- [xlwings 설치 및 운영체제별 의존성](https://docs.xlwings.org/en/stable/installation.html)
- [PyInstaller macOS 앱 빌드](https://pyinstaller.org/en/stable/usage.html#building-macos-app-bundles)
