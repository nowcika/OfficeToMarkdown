# OfficeToMarkdown

Word · Excel · PowerPoint 에서 **현재 열려 있는 파일**을 Markdown 으로 변환해 클립보드에 복사하는 VSTO 추가 기능.
각 앱의 **홈 탭 → Markdown → MD로 복사** 버튼을 누르면 변환해 클립보드에 넣는다. 파일은 저장하지 않는다.
클립보드에는 두 형식이 함께 들어가며, 붙여넣는 프로그램이 형식을 고른다.
- 메모장·코드 편집기 등: Markdown 텍스트 (이미지는 `![설명](<이름>_files/...)` 참조만 남고 파일은 없다)
- Word·Notion·메일 등 서식 편집기: HTML (제목·표·서식 유지, 이미지는 base64 로 내장)

### 복사하면 다른 프로그램이 자동으로 반응하게 하기

클립보드에는 `OfficeToMarkdown` 형식의 표식(`ClipboardMarker`: 복사마다 새 `id`, 원본 파일 이름)도 함께 들어간다.
클립보드를 감시하는 프로그램은 이 형식이 있을 때만 반응하면 다른 앱의 복사와 구분할 수 있다.

수신기: `samples/ClipboardListener` (Windows 전용, .NET 8)

```bat
dotnet run --project samples\ClipboardListener
```

창 없이 알림 영역(트레이)에서 실행되며, 추가 기능이 복사할 때마다 HTML 을 곧바로
`<저장 폴더>\<원본 이름>_<yyyyMMdd-HHmmss>.html` 로 저장하고 풍선 알림을 띄운다 (클릭하면 파일을 연다).
저장 폴더 기본값은 `문서\OfficeToMarkdown` 이며, **저장 폴더 변경...** 으로 고른 폴더는 `%AppData%\OfficeToMarkdown\ClipboardListener.txt` 에 기록되어 다음 실행에도 유지된다.
트레이 아이콘 오른쪽 클릭 메뉴: 저장 폴더 열기 · 저장 폴더 변경... · 마지막 파일 열기 · 종료. 두 번 실행해도 하나만 뜬다.

동작: `AddClipboardFormatListener` 로 `WM_CLIPBOARDUPDATE` 를 받아 표식을 확인하고, CF_HTML 에서 본문 조각을 꺼내
(`ClipboardHtml.ExtractFragment`) 완전한 HTML 문서로 저장한다. 한 번의 복사에 알림이 여러 번 올 수 있어 `id` 로 중복을 거른다.
Windows 시작 시 자동 실행하려면 빌드한 `ClipboardListener.exe` 의 바로 가기를 `shell:startup` 폴더에 넣는다.

## 구조

```
src/
  OfficeToMarkdown.Core/        netstandard2.0, Office 의존성 없음 (Linux/CI 에서도 빌드·테스트 가능)
    Model/                      블록 모델 (제목, 문단, 목록, 표, 인용, 구분선, 인라인 서식)
    MarkdownRenderer.cs         모델 → GFM 문자열 (이스케이프, 목록 번호, 표)
    HtmlRenderer.cs             모델 → HTML 조각 (클립보드 붙여넣기용, 이미지 data URI)
    HtmlPage.cs                 HTML 조각 → 저장용 완전한 HTML 문서 (표 스크롤·스타일)
    ClipboardHtml.cs            HTML 조각 → 클립보드 CF_HTML 형식
    ClipboardMarker.cs          클립보드 감시 프로그램용 표식 형식
    Word/WordOpenXmlConverter.cs   Document.WordOpenXML(Flat OPC) → 모델
    Excel/ExcelSheetConverter.cs   Range.Value 2차원 배열 → 표
  AddIns/
    Shared/ExportCommand.cs     리본 XML, 클립보드 복사, 오류 처리 (세 프로젝트에 "링크로 추가")
    WordToMarkdown/             ThisAddIn.cs, Ribbon.cs
    ExcelToMarkdown/            ThisAddIn.cs, Ribbon.cs
    PowerPointToMarkdown/       ThisAddIn.cs, Ribbon.cs, PresentationExtractor.cs
samples/ClipboardListener/            트레이 수신기: 복사되면 HTML 로 저장 (Windows 전용, 솔루션에는 포함하지 않음)
tests/OfficeToMarkdown.Core.Tests/   xUnit (dotnet test)
```

## 변환 규칙

| 앱 | 변환 내용 |
| --- | --- |
| Word | 제목 1~6(스타일·개요 수준, 한글 Word 포함), 굵게/기울임/취소선, 하이퍼링크, 글머리·번호 목록(중첩), 표(병합 셀 처리), 이미지, 콘텐츠 컨트롤, 변경 내용 추적(삭제분 제외) |
| Excel | 보이는 시트마다 `## 시트명` + 사용 범위 표 + 차트(PNG 이미지, 표 아래 위→아래 순). 차트 시트도 포함. 첫 행이 머리글, 날짜는 `yyyy-MM-dd`, 오류 값은 `#DIV/0!` 등으로 표시 |
| PowerPoint | 슬라이드마다 `## 번호. 제목`, 도형을 위→아래·왼쪽→오른쪽 순으로 읽음. 텍스트(글머리 수준·굵게·기울임·링크), 표, 그룹, SmartArt, 그림·차트·SVG 아이콘(PNG 로 렌더링), 발표자 노트(인용문). 슬라이드 사이 `---` |

**현재 제외:** PowerPoint 동영상·오디오, Excel 서식(셀 색, 백분율 표시 형식 등)·그림·도형, Word 각주·머리글/바닥글·텍스트 상자.
EMF/WMF 이미지는 원본 형식 그대로 넣으므로 대부분의 프로그램에서 표시되지 않는다.
PowerPoint 그림은 문서화되지 않은 `Shape.Export` 로 렌더링하므로 원본 파일이 아니라 자르기·크기 조절이 반영된 화면 모습이 들어간다.

## Visual Studio 에서 추가 기능 프로젝트 만들기 (Windows)

VSTO 프로젝트는 서명 인증서·매니페스트가 템플릿으로 생성되므로 Visual Studio 에서 만든 뒤 이 저장소의 파일을 넣는다.

준비물: Visual Studio 2022 + **Office/SharePoint 개발** 워크로드, Office 2016 이상(데스크톱).

1. `OfficeToMarkdown.sln` 을 연다.
2. 솔루션에 새 프로젝트 3개를 추가한다. **프로젝트 이름을 정확히 맞춰야** `ThisAddIn` 의 네임스페이스가 일치한다.
   - `Word VSTO 추가 기능` → `WordToMarkdown`
   - `Excel VSTO 추가 기능` → `ExcelToMarkdown`
   - `PowerPoint VSTO 추가 기능` → `PowerPointToMarkdown`
   - 위치: `src\AddIns\` (이미 있는 폴더와 같은 경로가 되도록), 프레임워크: .NET Framework 4.8
3. 각 프로젝트에서:
   1. 템플릿이 만든 `ThisAddIn.cs` 를 이 저장소의 같은 폴더 파일로 덮어쓴다.
      (`ThisAddIn.Designer.cs`, `ThisAddIn.Designer.xml` 은 그대로 둔다.)
   2. `Ribbon.cs` (PowerPoint 는 `PresentationExtractor.cs` 도) 를 **추가 → 기존 항목** 으로 넣는다.
   3. `src\AddIns\Shared\ExportCommand.cs` 를 **추가 → 기존 항목 → "추가" 버튼 옆 ▼ → 링크로 추가** 로 넣는다.
   4. **참조 추가**: 프로젝트 참조 `OfficeToMarkdown.Core`, 어셈블리 `System.Windows.Forms`.
4. 시작 프로젝트로 지정하고 F5 → Office 가 실행되며 홈 탭에 **MD로 복사** 버튼이 나타난다.

배포는 각 프로젝트의 **게시**(ClickOnce) 또는 MSI 설치 관리자를 사용한다.

## 테스트

```bash
dotnet test
```

Core 의 변환 로직(Word Flat OPC 파싱, Excel 값 서식, Markdown 렌더링)을 Office 없이 검증한다.
COM 에 의존하는 부분(`Ribbon.cs`, `PresentationExtractor.cs`)은 실제 Office 에서 F5 로 확인해야 한다.
