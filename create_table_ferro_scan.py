from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import os
import win32com.client as win32
import re

file_root = "C:\\Users\\pjy28\\Desktop\\develop"

root = Tk()  # 이미지선택창 열기
image_list = askopenfilenames()
root.destroy()  # 이미지선택창 닫기


# BASE_DIR = image_list[0]  # 이미지리스트에서 경로 추출

def extract_numbers(image_list):
    extracted_numbers = []
    for image_name in image_list:
        numbers = re.findall(r'\d+', image_name)
        last_three = numbers[-3:]
        last_three_numbers= ''.join(last_three)
        extracted_numbers.append(last_three_numbers)

    return extracted_numbers

# 주어진 image_list에서 숫자만 추출
name_number = extract_numbers(image_list)

sorted_numbers = sorted(name_number) # 숫자 오름차순


######################## 한글 시작 ##################################
hwp=win32.Dispatch("HWPFrame.HwpObject")

hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.XHwpWindows.Item(0).Visible = True


# 가운데 정렬
def align_center():
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.Run("ParagraphShapeAlignCenter")
    hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HSecDef.HSet)

#  여백조정
def set_margin(left, right, top, bottom, header, footer) :

    hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
    # hwp.HParameterSet.HSecDef.PageDef.Landscape = 1  # 가로로
    hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(left)
    hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(right)
    hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(top)
    hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(bottom)
    hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(header)
    hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(footer)
    hwp.HParameterSet.HSecDef.PageDef.GutterLen = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 문서 전체 변경
    hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)

def create_table(row, column, table_width, table_height, image_url, image_idx, image_width, image_height): ##(줄 갯수, 칸 갯수, 사진 너비, 사진 높이, 사진 주소, 사진 번호)
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = row
    hwp.HParameterSet.HTableCreation.Cols = column
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HParameterSet.HTableCreation.HeightType = 1
    hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(table_width)
    hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(table_height)
    hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", column)
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(47)) ## 첫째줄 너비
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(25)) ## 둘째줄 너비
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(33)) ## 셋째줄 너비
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, hwp.MiliToHwpUnit(42)) ## 넷째줄 너비

    hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", row)
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(8)) ## 제목 높이
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(83)) ## 사진 높이
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2, hwp.MiliToHwpUnit(8)) ## 개요 높이
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(3, hwp.MiliToHwpUnit(10)) ## 내용 높이

    hwp.HParameterSet.HTableCreation.TableProperties.Width = mm_to_hu(table_width)
    hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginLeft = hwp.MiliToHwpUnit(0)
    hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginRight = hwp.MiliToHwpUnit(0)
    hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginTop = hwp.MiliToHwpUnit(0)
    hwp.HParameterSet.HTableCreation.TableProperties.OutsideMarginBottom = hwp.MiliToHwpUnit(0)
    hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    ###표 생성 완료###

    hwp.HAction.Run("TableCellBlockRow")  # 표 1행 선택
    hwp.HAction.Run("TableMergeCell")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "상록초등학교 1동교사 및 순성초등학교 1동교사 정밀안전점검 용역_순성초등학교 1동교사" ## 표 제목
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
    hwp.HAction.Run("TableCellBlockRow")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    hwp.HAction.Run("CharShapeBold")
    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("MoveDown")

    hwp.HAction.Run("TableCellBlockRow")  # 표 2행 선택
    hwp.HAction.Run("TableMergeCell")

 
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableCellBlockExtend")
    hwp.HAction.Run("TableCellBlockExtend")
    
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
	
    ##제목 글꼴 설정##
    hwp.HParameterSet.HCharShape.FaceNameUser = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameSymbol = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameOther = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameJapanese = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameHanja = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameLatin = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.FaceNameHangul = "신명조"
    hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType("HFT")
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(9.0)

    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);

    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveUp")


    ## 표 안에 이미지 넣기
    insert_image(image_width, image_height, image_url, image_idx)


    align_center()
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("TableColBegin")  # 열 맨 앞으로 이동

    ## 개요 작성
    insert_text(f'NO. {image_idx + 1}')
    hwp.HAction.Run("MoveRight")
    insert_text('부재명')
    hwp.HAction.Run("MoveRight")
    insert_text('설계배근')
    hwp.HAction.Run("MoveRight")
    insert_text('탐사배근')

    hwp.HAction.Run("TableCellBlockRow")
    align_center()
    hwp.HAction.Run("CharShapeBold")

    ##표안에 셀 색 변경하기
    act = hwp.CreateAction("CellFill")
    createSet = act.CreateSet()
    act.GetDefault(createSet)
    fillAttrSet = createSet.CreateItemSet("FillAttr", "DrawFillAttr")

    fillAttrSet.SetItem("Type", 1);
    fillAttrSet.SetItem("WinBrushFaceStyle", 0xffffffff);
    fillAttrSet.SetItem("WinBrushHatchColor", 0x00000000);
    fillAttrSet.SetItem("WinBrushFaceColor", hwp.RGBColor(229, 229, 229)); ## 표 셀 색상
    act.Execute(createSet);

    ##표 테두리 두줄로 변경(아래)
    hwp.HAction.GetDefault("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HParameterSet.HCellBorderFill.BorderTypeBottom = hwp.HwpLineType("DoubleSlim") ## Line Type 2중선
    hwp.HParameterSet.HCellBorderFill.BorderWidthBottom = hwp.HwpLineWidth("0.7mm") ## 굵기 0.7mm 되나?
    hwp.HAction.Execute("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
    hwp.HAction.Run("Cancel")

    ##표 밖으로 나오기
    hwp.HAction.Run("CloseEx")
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("BreakPara")



def mm_to_hu(value):
    return hwp.MiliToHwpUnit(value)

def insert_text(text): 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    ##사진 넣기
def insert_image(image_width, image_height, image_url, image_idx):
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
    hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{image_idx}"
    hwp.InsertPicture(image_url, True, 3, False, False, 0, image_width, image_height) 






####################실행####################################

align_center()
set_margin(15, 15, 20, 15, 15, 15)

for idx, image in enumerate(image_list):
    image_url = image_list[idx] # 이미지리스트에서 경로 추출
    create_table(4, 4, 250, 210, image_url, idx, 82, 82)