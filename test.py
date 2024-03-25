import win32com.client as win32

file_root = "C:\\Users\\pjy28\\Desktop\\develop"

hwp=win32.Dispatch("HWPFrame.HwpObject")

hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.XHwpWindows.Item(0).Visible = True

hwp.Open("C:\\Users\\pjy28\\Desktop\\develop\\cover_test.hwp")

hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option=hwp.HParameterSet.HFindReplace

constructionName =  option.FindString = "공사이름"
changedName = option.ReplaceString = "㈜두원 아산공장 신축공사"

option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)


hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option=hwp.HParameterSet.HFindReplace

constructionType =  option.FindString = "공사종류"
changedType = option.ReplaceString = "항타 및 항발기 1차"

option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)


hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option=hwp.HParameterSet.HFindReplace

constructionCompany =  option.FindString = "건설사"
changedCompany = option.ReplaceString = "항타 및 항발기 1차"

option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

hwp.SaveAs("C:\\Users\\pjy28\\Desktop\\develop\\1.표지.hwp")

new_filename = "1.표지.hwp"
new_file_path = file_root + "\\" + new_filename
hwp.SaveAs(new_file_path)


print('표지 템플릿 실행완료.')

