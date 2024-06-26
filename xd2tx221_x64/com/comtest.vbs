' xd2txcom.dll (xdoc2txt com dll版)サンプル
'===============================================================================
' xd2txcom.dllをregsvr32で登録してから実行してください。
'	regsvr32 xd2txcom.dll
'===============================================================================

'クラスオブジェクトを取得する
Set obj = CreateObject("xd2txcom.Xdoc2txt.1")

' 文書のテキストを抽出する
Dim fileText
fileText = obj.ExtractText("sample.doc",False)

MsgBox fileText

