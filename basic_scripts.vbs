'Author:  Birgit Kellner
'email: birgit.kellner@univie.ac.at 
Sub toTAT_full
'Andy says that sometime in the future these may have to be Variant types  ram to work with Array()
  Dim numbered(5) As String, accented(5) As String
  Dim n as long
  Dim oDoc as object, oReplace as object
  numbered() = Array("À","à","Á","á","Â","â","Ã","ã","Ä","ä","Å","å","Æ","æ","Ç","ç","È","è","É","é","Ê","ê","Ë","ë","Ì","ì","Í","í","Î","î","Ï","ï","Ð","ð","Ñ","ñ","Ò","ò","Ó","ó","Ô","ô","Õ","õ","Ö","ö","×","÷","Ø","ø","ú","û","Ü","ü","Ý","ý","Þ","þ","ß","ÿ","Û","¨","¸","¯","¿","ª","º","«","»","‰","¢","³","²","Ú","ù")
  accented() = Array("А","а","Б","б","В","в","Г","г","Д","д","Е","е","Ж","ж","З","з","И","и","Й","й","К","к","Л","л","М","м","Н","н","О","о","П","п","Р","р","С","с","Т","т","У","у","Ф","ф","Х","х","Ц","ц","Ч","ч","Ш","ш","ъ","ы","Ь","ь","Э","э","Ю","ю","Я","я","Ы", "Ә","ә","Ө","ө","Ү","ү","Ӊ","ң","Җ","җ","һ","Һ","Ъ","щ")
  oReplace = ThisComponent.createReplaceDescriptor()
  oReplace.SearchCaseSensitive = True
  For n = LBound(numbered()) To UBound(accented())
    oReplace.SearchString = numbered(n)
    oReplace.ReplaceString = accented(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
End Sub

'Author:  Birgit Kellner
'email: birgit.kellner@univie.ac.at 

Sub toTAT_short
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

'Author:  Birgit Kellner
'email: birgit.kellner@univie.ac.at 
'Andy says that sometime in the future these may have to be Variant types  ram to work with Array()
  Dim numbered(5) As String, accented(5) As String
  Dim n as long
  Dim oDoc as object, oReplace as object
  numbered() = Array("Ё","ё","Ї","ї","Є","є",">>","»","‰","ў","і","І")
  accented() = Array("Ә","ә","Ө","ө","Ү","ү","Ӊ","ӊ","Җ","җ","h","h")
  oReplace = ThisComponent.createReplaceDescriptor()
  oReplace.SearchCaseSensitive = True
  For n = LBound(numbered()) To UBound(accented())
    oReplace.SearchString = numbered(n)
    oReplace.ReplaceString = accented(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
End Sub

Sub toTAT_roz

rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

'Author:  Birgit Kellner
'email: birgit.kellner@univie.ac.at 
'Andy says that sometime in the future these may have to be Variant types  ram to work with Array()
  Dim numbered(5) As String, accented(5) As String
  Dim n as long
  Dim oDoc as object, oReplace as object
  numbered() = Array("¨","¸","Ї","¿","Є","º",">>","»","‰","¢","³","²")
  accented() = Array("Ә","ә","Ө","ө","Ү","ү","Ӊ","ӊ","Җ","җ","h","h")
  oReplace = ThisComponent.createReplaceDescriptor()
  oReplace.SearchCaseSensitive = True
  For n = LBound(numbered()) To UBound(accented())
    oReplace.SearchString = numbered(n)
    oReplace.ReplaceString = accented(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
End Sub

Sub toTAT_antiroz

rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

'Author:  Birgit Kellner
'email: birgit.kellner@univie.ac.at 
'Andy says that sometime in the future these may have to be Variant types  ram to work with Array()
  Dim numbered(5) As String, accented(5) As String
  Dim n as long
  Dim oDoc as object, oReplace as object
  numbered() = Array("Ә","ә","Ө","ө","Ү","ү","Ӊ","ӊ","Җ","җ","h","h")
  accented() = Array("¨","¸","Ї","¿","Є","º",">>","»","‰","¢","³","²")
  oReplace = ThisComponent.createReplaceDescriptor()
  oReplace.SearchCaseSensitive = True
  For n = LBound(numbered()) To UBound(accented())
    oReplace.SearchString = numbered(n)
    oReplace.ReplaceString = accented(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
End Sub

sub RemoveEmptyParsWorker
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(18) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.Backward"
args1(4).Value = false
args1(5).Name = "SearchItem.Pattern"
args1(5).Value = false
args1(6).Name = "SearchItem.Content"
args1(6).Value = false
args1(7).Name = "SearchItem.AsianOptions"
args1(7).Value = false
args1(8).Name = "SearchItem.AlgorithmType"
args1(8).Value = 1
args1(9).Name = "SearchItem.SearchFlags"
args1(9).Value = 65536
args1(10).Name = "SearchItem.SearchString"
args1(10).Value = "-$"
args1(11).Name = "SearchItem.ReplaceString"
args1(11).Value = "_СЛОВО_"
args1(12).Name = "SearchItem.Locale"
args1(12).Value = 255
args1(13).Name = "SearchItem.ChangedChars"
args1(13).Value = 2
args1(14).Name = "SearchItem.DeletedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.InsertedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.TransliterateFlags"
args1(16).Value = 1280
args1(17).Name = "SearchItem.Command"
args1(17).Value = 3
args1(18).Name = "Quiet"
args1(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())

rem ----------------------------------------------------------------------
dim args2(18) as new com.sun.star.beans.PropertyValue
args2(0).Name = "SearchItem.StyleFamily"
args2(0).Value = 2
args2(1).Name = "SearchItem.CellType"
args2(1).Value = 0
args2(2).Name = "SearchItem.RowDirection"
args2(2).Value = true
args2(3).Name = "SearchItem.AllTables"
args2(3).Value = false
args2(4).Name = "SearchItem.Backward"
args2(4).Value = false
args2(5).Name = "SearchItem.Pattern"
args2(5).Value = false
args2(6).Name = "SearchItem.Content"
args2(6).Value = false
args2(7).Name = "SearchItem.AsianOptions"
args2(7).Value = false
args2(8).Name = "SearchItem.AlgorithmType"
args2(8).Value = 1
args2(9).Name = "SearchItem.SearchFlags"
args2(9).Value = 65536
args2(10).Name = "SearchItem.SearchString"
args2(10).Value = "\.$"
args2(11).Name = "SearchItem.ReplaceString"
args2(11).Value = "_АБЗАЦ_"
args2(12).Name = "SearchItem.Locale"
args2(12).Value = 255
args2(13).Name = "SearchItem.ChangedChars"
args2(13).Value = 2
args2(14).Name = "SearchItem.DeletedChars"
args2(14).Value = 2
args2(15).Name = "SearchItem.InsertedChars"
args2(15).Value = 2
args2(16).Name = "SearchItem.TransliterateFlags"
args2(16).Value = 1280
args2(17).Name = "SearchItem.Command"
args2(17).Value = 3
args2(18).Name = "Quiet"
args2(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())

rem ----------------------------------------------------------------------
dim args3(18) as new com.sun.star.beans.PropertyValue
args3(0).Name = "SearchItem.StyleFamily"
args3(0).Value = 2
args3(1).Name = "SearchItem.CellType"
args3(1).Value = 0
args3(2).Name = "SearchItem.RowDirection"
args3(2).Value = true
args3(3).Name = "SearchItem.AllTables"
args3(3).Value = false
args3(4).Name = "SearchItem.Backward"
args3(4).Value = false
args3(5).Name = "SearchItem.Pattern"
args3(5).Value = false
args3(6).Name = "SearchItem.Content"
args3(6).Value = false
args3(7).Name = "SearchItem.AsianOptions"
args3(7).Value = false
args3(8).Name = "SearchItem.AlgorithmType"
args3(8).Value = 1
args3(9).Name = "SearchItem.SearchFlags"
args3(9).Value = 65536
args3(10).Name = "SearchItem.SearchString"
args3(10).Value = "$"
args3(11).Name = "SearchItem.ReplaceString"
args3(11).Value = " "
args3(12).Name = "SearchItem.Locale"
args3(12).Value = 255
args3(13).Name = "SearchItem.ChangedChars"
args3(13).Value = 2
args3(14).Name = "SearchItem.DeletedChars"
args3(14).Value = 2
args3(15).Name = "SearchItem.InsertedChars"
args3(15).Value = 2
args3(16).Name = "SearchItem.TransliterateFlags"
args3(16).Value = 1280
args3(17).Name = "SearchItem.Command"
args3(17).Value = 2
args3(18).Name = "Quiet"
args3(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args3())

rem ----------------------------------------------------------------------
dim args4(18) as new com.sun.star.beans.PropertyValue
args4(0).Name = "SearchItem.StyleFamily"
args4(0).Value = 2
args4(1).Name = "SearchItem.CellType"
args4(1).Value = 0
args4(2).Name = "SearchItem.RowDirection"
args4(2).Value = true
args4(3).Name = "SearchItem.AllTables"
args4(3).Value = false
args4(4).Name = "SearchItem.Backward"
args4(4).Value = false
args4(5).Name = "SearchItem.Pattern"
args4(5).Value = false
args4(6).Name = "SearchItem.Content"
args4(6).Value = false
args4(7).Name = "SearchItem.AsianOptions"
args4(7).Value = false
args4(8).Name = "SearchItem.AlgorithmType"
args4(8).Value = 1
args4(9).Name = "SearchItem.SearchFlags"
args4(9).Value = 65536
args4(10).Name = "SearchItem.SearchString"
args4(10).Value = "$"
args4(11).Name = "SearchItem.ReplaceString"
args4(11).Value = " "
args4(12).Name = "SearchItem.Locale"
args4(12).Value = 255
args4(13).Name = "SearchItem.ChangedChars"
args4(13).Value = 2
args4(14).Name = "SearchItem.DeletedChars"
args4(14).Value = 2
args4(15).Name = "SearchItem.InsertedChars"
args4(15).Value = 2
args4(16).Name = "SearchItem.TransliterateFlags"
args4(16).Value = 1280
args4(17).Name = "SearchItem.Command"
args4(17).Value = 3
args4(18).Name = "Quiet"
args4(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args4())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args6(18) as new com.sun.star.beans.PropertyValue
args6(0).Name = "SearchItem.StyleFamily"
args6(0).Value = 2
args6(1).Name = "SearchItem.CellType"
args6(1).Value = 0
args6(2).Name = "SearchItem.RowDirection"
args6(2).Value = true
args6(3).Name = "SearchItem.AllTables"
args6(3).Value = false
args6(4).Name = "SearchItem.Backward"
args6(4).Value = false
args6(5).Name = "SearchItem.Pattern"
args6(5).Value = false
args6(6).Name = "SearchItem.Content"
args6(6).Value = false
args6(7).Name = "SearchItem.AsianOptions"
args6(7).Value = false
args6(8).Name = "SearchItem.AlgorithmType"
args6(8).Value = 1
args6(9).Name = "SearchItem.SearchFlags"
args6(9).Value = 65536
args6(10).Name = "SearchItem.SearchString"
args6(10).Value = "_СЛОВО_ "
args6(11).Name = "SearchItem.ReplaceString"
args6(11).Value = ""
args6(12).Name = "SearchItem.Locale"
args6(12).Value = 255
args6(13).Name = "SearchItem.ChangedChars"
args6(13).Value = 2
args6(14).Name = "SearchItem.DeletedChars"
args6(14).Value = 2
args6(15).Name = "SearchItem.InsertedChars"
args6(15).Value = 2
args6(16).Name = "SearchItem.TransliterateFlags"
args6(16).Value = 1280
args6(17).Name = "SearchItem.Command"
args6(17).Value = 3
args6(18).Name = "Quiet"
args6(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(0) as new com.sun.star.beans.PropertyValue
args7(0).Name = "ControlCodes"
args7(0).Value = false

dispatcher.executeDispatch(document, ".uno:ControlCodes", "", 0, args7())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args9(18) as new com.sun.star.beans.PropertyValue
args9(0).Name = "SearchItem.StyleFamily"
args9(0).Value = 2
args9(1).Name = "SearchItem.CellType"
args9(1).Value = 0
args9(2).Name = "SearchItem.RowDirection"
args9(2).Value = true
args9(3).Name = "SearchItem.AllTables"
args9(3).Value = false
args9(4).Name = "SearchItem.Backward"
args9(4).Value = false
args9(5).Name = "SearchItem.Pattern"
args9(5).Value = false
args9(6).Name = "SearchItem.Content"
args9(6).Value = false
args9(7).Name = "SearchItem.AsianOptions"
args9(7).Value = false
args9(8).Name = "SearchItem.AlgorithmType"
args9(8).Value = 1
args9(9).Name = "SearchItem.SearchFlags"
args9(9).Value = 65536
args9(10).Name = "SearchItem.SearchString"
args9(10).Value = "_АБЗАЦ_ "
args9(11).Name = "SearchItem.ReplaceString"
args9(11).Value = ".\n"
args9(12).Name = "SearchItem.Locale"
args9(12).Value = 255
args9(13).Name = "SearchItem.ChangedChars"
args9(13).Value = 2
args9(14).Name = "SearchItem.DeletedChars"
args9(14).Value = 2
args9(15).Name = "SearchItem.InsertedChars"
args9(15).Value = 2
args9(16).Name = "SearchItem.TransliterateFlags"
args9(16).Value = 1280
args9(17).Name = "SearchItem.Command"
args9(17).Value = 3
args9(18).Name = "Quiet"
args9(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args9())

rem ----------------------------------------------------------------------
dim args10(18) as new com.sun.star.beans.PropertyValue
args10(0).Name = "SearchItem.StyleFamily"
args10(0).Value = 2
args10(1).Name = "SearchItem.CellType"
args10(1).Value = 0
args10(2).Name = "SearchItem.RowDirection"
args10(2).Value = true
args10(3).Name = "SearchItem.AllTables"
args10(3).Value = false
args10(4).Name = "SearchItem.Backward"
args10(4).Value = false
args10(5).Name = "SearchItem.Pattern"
args10(5).Value = false
args10(6).Name = "SearchItem.Content"
args10(6).Value = false
args10(7).Name = "SearchItem.AsianOptions"
args10(7).Value = false
args10(8).Name = "SearchItem.AlgorithmType"
args10(8).Value = 1
args10(9).Name = "SearchItem.SearchFlags"
args10(9).Value = 65536
args10(10).Name = "SearchItem.SearchString"
args10(10).Value = "_АБЗАЦ_"
args10(11).Name = "SearchItem.ReplaceString"
args10(11).Value = ".\n"
args10(12).Name = "SearchItem.Locale"
args10(12).Value = 255
args10(13).Name = "SearchItem.ChangedChars"
args10(13).Value = 2
args10(14).Name = "SearchItem.DeletedChars"
args10(14).Value = 2
args10(15).Name = "SearchItem.InsertedChars"
args10(15).Value = 2
args10(16).Name = "SearchItem.TransliterateFlags"
args10(16).Value = 1280
args10(17).Name = "SearchItem.Command"
args10(17).Value = 3
args10(18).Name = "Quiet"
args10(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args10())


end sub

Sub toTAT_full_programma_vicecersa
'Andy says that sometime in the future these may have to be Variant types  ram to work with Array()
  Dim numbered(5) As String, accented(5) As String, quotes_from(2) As String, quotes_to(2) As String
  Dim n as long
  Dim oDoc as object, oReplace as object
  quotes_from = Array("«", "»")
  quotes_to = Array("“", "”")
  accented() = Array("À","à","Á","á","Â","â","Ã","ã","Ä","ä","Å","å","Æ","æ","Ç","ç","È","è","É","é","Ê","ê","Ë","ë","Ì","ì","Í","í","Î","î","Ï","ï","Ð","ð","Ñ","ñ","Ò","ò","Ó","ó","Ô","ô","Õ","õ","Ö","ö","×","÷","Ø","ø","ú","û","Ü","ü","Ý","ý","Þ","þ","ß","ÿ","Û","¨","¸","¯","¿","ª","º","«","»","‰","¢","³","²","Ú","ù")
  numbered() = Array("А","а","Б","б","В","в","Г","г","Д","д","Е","е","Ж","ж","З","з","И","и","Й","й","К","к","Л","л","М","м","Н","н","О","о","П","п","Р","р","С","с","Т","т","У","у","Ф","ф","Х","х","Ц","ц","Ч","ч","Ш","ш","ъ","ы","Ь","ь","Э","э","Ю","ю","Я","я","Ы", "Ә","ә","Ө","ө","Ү","ү","Ӊ","ң","Җ","җ","һ","Һ","Ъ","щ")
  oReplace = ThisComponent.createReplaceDescriptor()
  oReplace.SearchCaseSensitive = True
  For n = LBound(quotes_from()) To UBound(quotes_to())
    oReplace.SearchString = quotes_from(n)
    oReplace.ReplaceString = quotes_to(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
  For n = LBound(numbered()) To UBound(accented())
    oReplace.SearchString = numbered(n)
    oReplace.ReplaceString = accented(n)
    ThisComponent.ReplaceAll(oReplace)
  Next n
End Sub
