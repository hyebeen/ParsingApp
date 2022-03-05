# -*- coding: utf-8 -*-
# PYTHON 3.6, 3.7, 3.9

# pyinstaller --icon=pngegg.ico --onefile main.py

# 개선할 점
# 1. utf-8 변환 에러 수정
# 2. 페이지 나누기 미리보기

import io, os, time, sys, re
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.simplefilter("ignore")

def ConvertFile(pathScriptDir, fileList) :
    for fileName in fileList:
        fileDir = pathScriptDir + '\\' + fileName
        try:
            file = io.open(fileDir, mode='r').read()
        except UnicodeError:
            continue
        io.open(fileDir, mode='w', encoding="utf-8").write(file)


def InputInformation():
    menu = input("* 파싱 결과 양식을 선택해주세요.\n  하나의 시트로 생성:1     여러 시트로 생성:2\n  => ")

    sampleReportFile = input("* 상세결과 파일 드래그 또는 입력 (파일 내 sample 시트가 존재해야 합니다.)\n  Ex) C:\\Users\hyebeen\\Desktop\\01_OOO_시스템 취약점 진단 결과_Unix_v0.1.xlsx \n  => ")
    sampleReportFile = sampleReportFile.strip('"')

    pathScriptDir = input("* 스크립트 결과 파일이 있는 폴더 드래그 또는 입력\n  Ex) C:\\Users\hyebeen\\Desktop\\directory \n  => ")
    try :
        fileList = os.listdir(pathScriptDir)
    except FileNotFoundError :
        print("* 잘못된 경로")
        time.sleep(2)
        sys.exit()

    startLine = input("* \"시작 패턴\"을 입력\n  (default) ##### START ##### \n  =>  ")
    if startLine == "" : startLine = "##### START #####"

    endLine = input("* \"마지막 패턴\"을 입력\n  (default) ##### END ##### \n  =>  ")
    if endLine == "" : endLine = "##### END #####"

    return pathScriptDir, sampleReportFile, fileList, startLine, endLine, menu


def RemoveFile() :
    if os.path.isfile("(파싱결과) 시스템 취약점 진단 결과.xlsx"):
        removeFlag = input("* 기존 파일 \"(파싱결과) 시스템 취약점 진단 결과.xlsx\"이 존재합니다.\n* 계속 진행하려면 삭제해야 합니다. 삭제하시겠습니까?\n  예:y, 아니오:n\n =>  ")
        if removeFlag == "y" :
            try:
                os.remove("(파싱결과) 시스템 취약점 진단 결과.xlsx")
            except PermissionError:
                print("* \"(파싱결과) 시스템 취약점 진단 결과.xlsx\"파일을 닫은 후 재실행해주세요.\n")
                time.sleep(2)
                sys.exit()
            print("* 기존 파일 \"(파싱결과) 시스템 취약점 진단 결과.xlsx\"을 삭제하였습니다.")
        if removeFlag == "n":
            print("기존 파일 \"(파싱결과) 시스템 취약점 진단 결과.xlsx\" 삭제 후 실행해야 합니다.\n")
            time.sleep(2)
            sys.exit()


def firstMenu(pathScriptDir, sampleReportFile, fileList, startLine, endLine) :

    wb = load_workbook(sampleReportFile)
    copy_ws = wb.copy_worksheet(wb["sample"])
    copy_ws.title = "파싱 결과"
    readFlag = "False"

    #sample 시트 수정해서 틀 만들기
    copy_ws.unmerge_cells('A1:A3')
    copy_ws.unmerge_cells('E1:F1')
    copy_ws.unmerge_cells('E2:F2')
    copy_ws.unmerge_cells('E3:F3')

    copy_ws.delete_rows(1, 4)
    copy_ws.row_dimensions[4].height = 18
    copy_ws.column_dimensions["I"].hidden=False


    rowIndex = 1
    colIndex = 7

    for fileName in fileList:
        fileDir = pathScriptDir + '\\' + fileName
        txt_f = open(fileDir, encoding="utf-8", mode="r")

        while True:
            txt = txt_f.readline()
            if startLine in txt :
                write_txt =""
                readFlag = "True"
                rowIndex += 1
                continue
            if endLine in txt :
                readFlag = "False"
                copy_ws.cell(rowIndex, 9, re.search('.+_[0-9]+.[0-9]+.[0-9]+.[0-9]+_(.+).log', fileName).group(1))
                copy_ws.cell(rowIndex, colIndex, write_txt)
                copy_ws.cell(rowIndex, colIndex).alignment = Alignment(wrap_text=True)
            if readFlag == "True":
                write_txt += txt
            if not txt: break

        txt_f.close()

    # 행 높이 설정
    for row in range(1, rowIndex + 1):
        copy_ws.row_dimensions[row].height = 18

    # 파일 저장
    wb.save(filename="(파싱결과) 시스템 취약점 진단 결과.xlsx")


def secondMenu(pathScriptDir, sampleReportFile, fileList, startLine, endLine) :
    wb = load_workbook(sampleReportFile)
    readFlag = "False"

    for fileName in fileList:
        fileDir = pathScriptDir + '\\' + fileName
        txt_f = open(fileDir, encoding="utf-8", mode="r")

        copy_ws = wb.copy_worksheet(wb["sample"])
        copy_ws.title = re.search('.*_([a-zA-Z0-9]+).log', fileName).group(1)
        copy_ws.sheet_view.showGridLines = False

        rowIndex = 5
        colIndex = 7

        while True:
            txt = txt_f.readline()
            if startLine in txt:
                write_txt = ""
                readFlag = "True"
                rowIndex += 1
                continue
            if endLine in txt:
                readFlag = "False"
                copy_ws.cell(rowIndex, colIndex, write_txt)
                copy_ws.cell(rowIndex, colIndex).alignment = Alignment(wrap_text=True)
            if readFlag == "True":
                write_txt += txt
            if not txt: break

    # 파일 저장
    wb.save(filename="(파싱결과) 시스템 취약점 진단 결과.xlsx")


def main():
    # 기존 파일 있으면 삭제
    RemoveFile()

    # 정보 입력
    pathScriptDir, sampleReportFile, fileList, startLine, endLine, menu = InputInformation()

    # utf-8로 변환
    ConvertFile(pathScriptDir, fileList)

    # 파싱 실행
    if menu == '1':
        firstMenu(pathScriptDir, sampleReportFile, fileList, startLine, endLine)
    elif menu == '2':
        secondMenu(pathScriptDir, sampleReportFile, fileList, startLine, endLine)

    # 종료
    print("*** [완료] \"파싱 프로그램\"이 있는 경로에 \"(파싱결과) 시스템 취약점 진단 결과.xlsx\" 결과 파일 생성 ***")
    time.sleep(5)


if __name__ == '__main__':
    main()
