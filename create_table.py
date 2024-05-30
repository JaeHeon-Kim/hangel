from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import os
import win32com.client as win32
import re

file_root = "C:\\Users\\pjy28\\Desktop\\develop"

root = Tk()  # 이미지선택창 열기
image_list = askopenfilenames()
root.destroy()  # 이미지선택창 닫기
BASE_DIR = image_list[0]  # 이미지리스트에서 경로 추출

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


# %% 표_리스트 만들기

# 표_리스트 = list(set([i.split("_")[0][:-1] for i in image_list]))
# 표_리스트.sort()

# print(표_리스트)

hwp=win32.Dispatch("HWPFrame.HwpObject")

hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.XHwpWindows.Item(0).Visible = True

# hwp.Open("C:\\Users\\pjy28\\Desktop\\develop\\picture_test.hwpx")

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

def create_table(row, column, width, height):
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = row
    hwp.HParameterSet.HTableCreation.Cols = column
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HParameterSet.HTableCreation.HeightType = 1
    hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(width)
    hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(height)
    hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", column)
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(47))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(25))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(33))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, hwp.MiliToHwpUnit(42))

    hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", row)
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(8))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(83))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2, hwp.MiliToHwpUnit(8))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(3, hwp.MiliToHwpUnit(11))
    # for i in range(column):
    #     hwp.HParameterSet.HTableCreation.ColWidth.item[i] = hwp.MiliToHwpUnit((width / column) - 3.6)
    # hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", row)
    # for i in range(row):
    #     hwp.HParameterSet.HTableCreation.RowHeight.item[i] = hwp.MiliToHwpUnit((height / row) - 1.0)
    hwp.HParameterSet.HTableCreation.TableProperties.Width = mm_to_hu(width)
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
    hwp.HParameterSet.HInsertText.Text = "상록초등학교 1동교사 및 순성초등학교 1동교사 정밀안전점검 용역_순성초등학교 1동교사"
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

    for idx, content in enumerate(sorted_numbers):
         
         hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
         hwp.HParameterSet.HShapeObject.ShapeTableCell.Editable = 1
         hwp.HParameterSet.HShapeObject.ShapeTableCell.CellCtrlData.name = f"{idx}"
    # InsertPicture(path, embeded, sizeoption, reverse, watermark,effect, width, height, callback)  
    hwp.InsertPicture(BASE_DIR, True, 3, False, False, 0, 82, 82) 

    align_center()
    
    hwp.HAction.Run("MoveDown")

    hwp.HAction.Run("TableColBegin")  # 열 맨 앞으로 이동

    insert_text('NO. 1')
    hwp.HAction.Run("MoveRight")
    insert_text('부재명')
    hwp.HAction.Run("MoveRight")
    insert_text('설계배근')
    hwp.HAction.Run("MoveRight")
    insert_text('탐사배근')

    hwp.HAction.Run("TableCellBlockRow")
    align_center()
    hwp.HAction.Run("CharShapeBold")

def mm_to_hu(value):
    return hwp.MiliToHwpUnit(value)

def insert_text(text): 
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

align_center()
set_margin(15, 15, 20, 15, 15, 15)
create_table(4, 4, 250, 210)