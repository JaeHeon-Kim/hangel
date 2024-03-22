import win32com.client as win32

hwp=win32.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
hwp.XHwpWindows.Item(0).Visible = True


hwp.Open(r"C:\Users\jaehe\OneDrive\바탕 화면\hangel\test.hwpx")


hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
option=hwp.HParameterSet.HFindReplace

option.FindString = "나는야"
option.ReplaceString = "아아아"
option.IgnoreMessage = 1
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

option.FindString = "슈퍼맨"
option.ReplaceString = "배트맨"
option.IgnoreMessage = 1

hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)