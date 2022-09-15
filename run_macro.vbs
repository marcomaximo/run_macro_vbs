Dim objExcel
Dim objWB


Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True


Set objWB = objExcel.Workbooks.Open("caminho_do_seu_arquivo")
WScript.Sleep 10000
objExcel.Run "nome_da_macro"


objWB.Save
WScript.Sleep 5000
objWB.Close True
objExcel.Quit