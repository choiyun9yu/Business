# MS ACCESS 

## 1. Access 기본 개요
### 1-1. What is Access?
- 소규모 자영업자 같은 비전문가도 데이터베이스를 다룰 수 있도록 만든 데이터베이스 프로그램
- 수많은 데이터를 등록하고 그들 사이의 관계설정, 조회, 출력 등을 간편하게 프로그램화 가능

### 1-2. What is different from Excel?
- (Excel) 데이터를 계산작업, 통계적 분석을 위한 도구
- (Access) 데이터를 구조화하여 체계적으로 관리하기 위한 도구
| 구분  | Excel       | Access       |
| --- | ----------- | ------------ |
| 성격  | 스프레드시트      | 관계형 데이터베이스   |
| 구조  | 셀 기반 (행·열)  | 테이블 + 관계     |
| 목적  | 계산, 분석, 시각화 | 데이터 저장·관리·검색 |
| 사용자 | 개인 분석 중심    | 조직/업무 데이터 관리 |

### 1-3. VBA(Visual Basic for Applications)
- Microsoft Office Program 안에서 자동화·개발을 할 수 있게 해주는 내장 프로그래밍 언어  
  ex) 버튼 클릭하면 자동 계산, 외부 API 호출해서 데이터 가져오기, 반복 업무 자동 처리, 보고서 자동 생성, DB 자동 저장 등
#### VBA 로 할 수 있는 것
| 기능      | 예시               |
| ------- | ---------------- |
| 자동화     | 매일 환율 자동 수집      |
| 데이터 처리  | 1만 건 반복 계산       |
| 외부 연결   | Open API 호출      |
| DB 제어   | Access 테이블 자동 입력 |
| 보고서 자동화 | 버튼 하나로 PDF 생성    |



## 2. Waht is Database?
- 서로 관련성있는 운영가능한 데이터들의 집합
- 관련성있는 데이터들을 저장하기 위한 저장소

### 2-1. Table
- 필드와 이를 구성하는 레코드 집함
- 테이블 생성시 각 필드의 데이터 타입도 반드시 고려
- 테이블은 데이터베이스 구축의 기본

### 2-3. What is Access?
- 사용자가 데이터베이스를 참조하고 관리하기 위한 데이터베이스 응용프로그램
- (폼) 데이터를 등록하거나 검색 작업을 하기 위한 화면
- (보고서) 데이터를 출력하기 위한 양식 화면



## 3. Access Database 작업
### 3-1. 서식 데이터베이스
#### Access 에서 기본적으로 제공하는 서식 활용 가능 
  <img width="2006" height="771" alt="image" src="https://github.com/user-attachments/assets/a2f3e037-44e4-4d96-af22-a23677f0559a" />
- 내가 쓰고자하는 기능과 비숫한 서식을 찾고 약간만 고쳐서 사용가능
- 기존 폼 [우클릭] - [디자인보기] 로 수정가능

### 3-1. 테이블 생성
#### 새 테이블 만들기
  <img width="1106" height="254" alt="image" src="https://github.com/user-attachments/assets/ee57f0f8-fb67-4192-99b8-6a5cba84f73f" />

#### 필드 데이터 타입 정의 (다른 DMBS와 호환성을 위해 필드명은 영문으로 쓰는게 좋음)
<img width="337" height="451" alt="image" src="https://github.com/user-attachments/assets/67abbaa6-bf25-4bb1-abe1-285067d823e8" />
- 짧은 텍스트(255자 이내)
- 긴 텍스트(약 1GB 이내, 약 64,000자)
- 숫자
- 큰 숫자(8 Byte, SQLServer와 호환)
- 일련번호(Primary Key)
- OLE 개체(그림, 이미지, 그래픽 등 표시하기 위한 데이터 타입)

#### 필드속성 정의
<img width="1186" height="482" alt="image" src="https://github.com/user-attachments/assets/574a3eb5-7917-4895-9617-ce69907e9fa9" />
- 데이터 입력시 사용자가 원하는 데이터 크기, 조건 등을 지정하여 견고한 데이터 입력을 유도하는 것
- 데이터 타입 마다 필드 속성은 다를 수 있음
- 입력 마스크
  <img width="525" height="304" alt="image" src="https://github.com/user-attachments/assets/26f4c62c-416b-4040-bd9f-0d94d9f447da" />
- 


### 3-2. 관계 설정


### 3-3. 쿼리

### 3-4. 조인



## 4. 폼 작업
### 4-1. 폼 작성방법


### 4-2. 폼 속성 시트


### 4-3. 콤보박스


### 4-4. 하위폼 / 하위보고서


### 4-5. 폼에 차트 삽입




## 5. 보고서 작업
### 5-1. 보고서 생성


### 5-2. 보고서 그룹화





















# Excel 에서 VBA 사용하기
[파일] - [옵션] - [리본 사용자지정] - [리본 메뉴 사용자지정] - [개발도구] 체크  
[개발도구] - [Visual Basic]  
<br>
[프로젝트] - [현재_통합_문서]  
<br>
```
# 엑셀 팝업창에 안녕하세요. 출력
Sub Hello()
    MsgBox "안녕하세요."
End Sub

[F5] - [실행]
```

# Access 에서 VBA 사용하기
[데이터베이스 도구] - [Visual Basic]
