# CADExtractor
캐드 면적 추출 라이브러리

## 프로젝트 개요
- polyline으로 닫힌 부분의 면적을 추출
- 추출할 대상은 layer 이름으로 구별

## 개발 환경
- C#, Visual Studio 2019
- AutoCAD 2012
- acmgd.dll, acdbmgd.dll 참조

## 사용법
1. NETLOAD로 [CADExtractorLib.dll](https://github.com/lcw3176/CADExtractor/releases/download/v1.1.0/CADExtractorLib.dll) 임포트
2. EXTRACT 명령어 입력
3. 엑셀 경로 입력
4. 추출할 레이어 명 입력
```
Command: NETLOAD
Command: EXTRACT

## 확장자 입력 불필요
## 존재하는 파일명 or 새로 만들 파일명 입력 
## '/' 안됨, '\'(역슬래시) 로 입력해야함
Enter the Excel Path: D:\sample

## 전체 추출: *
## EX) Enter the Layer Name: *
##
## 다수의 레이어: 쉼표(,) 로 구분
## EX) Enter the Layer Name: L-BO, L-도로
##
## 단일 레이어: 이름 입력
## EX) Enter the Layer Name: L-BO

Enter the Layer Name: *
Writing Excel....
Complete
```

## 결과물 샘플
![화면 캡처 2021-10-10 165102](https://user-images.githubusercontent.com/59993347/136687332-31b82bfc-e855-4b83-81b2-a5f860e9ed77.jpg)
