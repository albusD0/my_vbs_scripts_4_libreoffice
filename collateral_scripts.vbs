REM  *****  BASIC  *****

sub Abzaz
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(21) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.SearchFiltered"
args1(4).Value = false
args1(5).Name = "SearchItem.Backward"
args1(5).Value = false
args1(6).Name = "SearchItem.Pattern"
args1(6).Value = false
args1(7).Name = "SearchItem.Content"
args1(7).Value = false
args1(8).Name = "SearchItem.AsianOptions"
args1(8).Value = false
args1(9).Name = "SearchItem.AlgorithmType"
args1(9).Value = 1
args1(10).Name = "SearchItem.SearchFlags"
args1(10).Value = 65536
args1(11).Name = "SearchItem.SearchString"
args1(11).Value = "\t"
args1(12).Name = "SearchItem.ReplaceString"
args1(12).Value = ""
args1(13).Name = "SearchItem.Locale"
args1(13).Value = 255
args1(14).Name = "SearchItem.ChangedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.DeletedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.InsertedChars"
args1(16).Value = 2
args1(17).Name = "SearchItem.TransliterateFlags"
args1(17).Value = 1280
args1(18).Name = "SearchItem.Command"
args1(18).Value = 3
args1(19).Name = "SearchItem.SearchFormatted"
args1(19).Value = false
args1(20).Name = "SearchItem.AlgorithmType2"
args1(20).Value = 2
args1(21).Name = "Quiet"
args1(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())

rem ----------------------------------------------------------------------
dim args2(21) as new com.sun.star.beans.PropertyValue
args2(0).Name = "SearchItem.StyleFamily"
args2(0).Value = 2
args2(1).Name = "SearchItem.CellType"
args2(1).Value = 0
args2(2).Name = "SearchItem.RowDirection"
args2(2).Value = true
args2(3).Name = "SearchItem.AllTables"
args2(3).Value = false
args2(4).Name = "SearchItem.SearchFiltered"
args2(4).Value = false
args2(5).Name = "SearchItem.Backward"
args2(5).Value = false
args2(6).Name = "SearchItem.Pattern"
args2(6).Value = false
args2(7).Name = "SearchItem.Content"
args2(7).Value = false
args2(8).Name = "SearchItem.AsianOptions"
args2(8).Value = false
args2(9).Name = "SearchItem.AlgorithmType"
args2(9).Value = 1
args2(10).Name = "SearchItem.SearchFlags"
args2(10).Value = 65536
args2(11).Name = "SearchItem.SearchString"
args2(11).Value = "\n"
args2(12).Name = "SearchItem.ReplaceString"
args2(12).Value = "\n"
args2(13).Name = "SearchItem.Locale"
args2(13).Value = 255
args2(14).Name = "SearchItem.ChangedChars"
args2(14).Value = 2
args2(15).Name = "SearchItem.DeletedChars"
args2(15).Value = 2
args2(16).Name = "SearchItem.InsertedChars"
args2(16).Value = 2
args2(17).Name = "SearchItem.TransliterateFlags"
args2(17).Value = 1280
args2(18).Name = "SearchItem.Command"
args2(18).Value = 3
args2(19).Name = "SearchItem.SearchFormatted"
args2(19).Value = false
args2(20).Name = "SearchItem.AlgorithmType2"
args2(20).Value = 2
args2(21).Name = "Quiet"
args2(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())

rem ----------------------------------------------------------------------
dim args5(21) as new com.sun.star.beans.PropertyValue
args5(0).Name = "SearchItem.StyleFamily"
args5(0).Value = 2
args5(1).Name = "SearchItem.CellType"
args5(1).Value = 0
args5(2).Name = "SearchItem.RowDirection"
args5(2).Value = true
args5(3).Name = "SearchItem.AllTables"
args5(3).Value = false
args5(4).Name = "SearchItem.SearchFiltered"
args5(4).Value = false
args5(5).Name = "SearchItem.Backward"
args5(5).Value = false
args5(6).Name = "SearchItem.Pattern"
args5(6).Value = false
args5(7).Name = "SearchItem.Content"
args5(7).Value = false
args5(8).Name = "SearchItem.AsianOptions"
args5(8).Value = false
args5(9).Name = "SearchItem.AlgorithmType"
args5(9).Value = 1
args5(10).Name = "SearchItem.SearchFlags"
args5(10).Value = 65536
args5(11).Name = "SearchItem.SearchString"
args5(11).Value = "  "
args5(12).Name = "SearchItem.ReplaceString"
args5(12).Value = ""
args5(13).Name = "SearchItem.Locale"
args5(13).Value = 255
args5(14).Name = "SearchItem.ChangedChars"
args5(14).Value = 2
args5(15).Name = "SearchItem.DeletedChars"
args5(15).Value = 2
args5(16).Name = "SearchItem.InsertedChars"
args5(16).Value = 2
args5(17).Name = "SearchItem.TransliterateFlags"
args5(17).Value = 1280
args5(18).Name = "SearchItem.Command"
args5(18).Value = 3
args5(19).Name = "SearchItem.SearchFormatted"
args5(19).Value = false
args5(20).Name = "SearchItem.AlgorithmType2"
args5(20).Value = 2
args5(21).Name = "Quiet"
args5(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args5())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, Array())


end sub

sub kavichki
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args2(23) as new com.sun.star.beans.PropertyValue
args2(0).Name = "SearchItem.StyleFamily"
args2(0).Value = 2
args2(1).Name = "SearchItem.CellType"
args2(1).Value = 0
args2(2).Name = "SearchItem.RowDirection"
args2(2).Value = true
args2(3).Name = "SearchItem.AllTables"
args2(3).Value = false
args2(4).Name = "SearchItem.SearchFiltered"
args2(4).Value = false
args2(5).Name = "SearchItem.Backward"
args2(5).Value = false
args2(6).Name = "SearchItem.Pattern"
args2(6).Value = false
args2(7).Name = "SearchItem.Content"
args2(7).Value = false
args2(8).Name = "SearchItem.AsianOptions"
args2(8).Value = false
args2(9).Name = "SearchItem.AlgorithmType"
args2(9).Value = 1
args2(10).Name = "SearchItem.SearchFlags"
args2(10).Value = 65536
args2(11).Name = "SearchItem.SearchString"
args2(11).Value = "“"
args2(12).Name = "SearchItem.ReplaceString"
args2(12).Value = CHR$(171)
args2(13).Name = "SearchItem.Locale"
args2(13).Value = 255
args2(14).Name = "SearchItem.ChangedChars"
args2(14).Value = 2
args2(15).Name = "SearchItem.DeletedChars"
args2(15).Value = 2
args2(16).Name = "SearchItem.InsertedChars"
args2(16).Value = 2
args2(17).Name = "SearchItem.TransliterateFlags"
args2(17).Value = 1280
args2(18).Name = "SearchItem.Command"
args2(18).Value = 3
args2(19).Name = "SearchItem.SearchFormatted"
args2(19).Value = false
args2(20).Name = "SearchItem.AlgorithmType2"
args2(20).Value = 2
args2(21).Name = "Quiet"
args2(21).Value = true
args2(22).Name = "SearchItem.SearchString"
args2(22).Value = "Гизатуллина"
args2(23).Name = "SearchItem.ReplaceString"
args2(23).Value = "Гиззатуллина"


dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args4(21) as new com.sun.star.beans.PropertyValue
args4(0).Name = "SearchItem.StyleFamily"
args4(0).Value = 2
args4(1).Name = "SearchItem.CellType"
args4(1).Value = 0
args4(2).Name = "SearchItem.RowDirection"
args4(2).Value = true
args4(3).Name = "SearchItem.AllTables"
args4(3).Value = false
args4(4).Name = "SearchItem.SearchFiltered"
args4(4).Value = false
args4(5).Name = "SearchItem.Backward"
args4(5).Value = false
args4(6).Name = "SearchItem.Pattern"
args4(6).Value = false
args4(7).Name = "SearchItem.Content"
args4(7).Value = false
args4(8).Name = "SearchItem.AsianOptions"
args4(8).Value = false
args4(9).Name = "SearchItem.AlgorithmType"
args4(9).Value = 1
args4(10).Name = "SearchItem.SearchFlags"
args4(10).Value = 65536
args4(11).Name = "SearchItem.SearchString"
args4(11).Value = "”"
args4(12).Name = "SearchItem.ReplaceString"
args4(12).Value = CHR$(187)
args4(13).Name = "SearchItem.Locale"
args4(13).Value = 255
args4(14).Name = "SearchItem.ChangedChars"
args4(14).Value = 2
args4(15).Name = "SearchItem.DeletedChars"
args4(15).Value = 2
args4(16).Name = "SearchItem.InsertedChars"
args4(16).Value = 2
args4(17).Name = "SearchItem.TransliterateFlags"
args4(17).Value = 1280
args4(18).Name = "SearchItem.Command"
args4(18).Value = 3
args4(19).Name = "SearchItem.SearchFormatted"
args4(19).Value = false
args4(20).Name = "SearchItem.AlgorithmType2"
args4(20).Value = 2
args4(21).Name = "Quiet"
args4(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args4())

REM  *****  BASIC  *****


rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args22(21) as new com.sun.star.beans.PropertyValue
args22(0).Name = "SearchItem.StyleFamily"
args22(0).Value = 2
args22(1).Name = "SearchItem.CellType"
args22(1).Value = 0
args22(2).Name = "SearchItem.RowDirection"
args22(2).Value = true
args22(3).Name = "SearchItem.AllTables"
args22(3).Value = false
args22(4).Name = "SearchItem.SearchFiltered"
args22(4).Value = false
args22(5).Name = "SearchItem.Backward"
args22(5).Value = false
args22(6).Name = "SearchItem.Pattern"
args22(6).Value = false
args22(7).Name = "SearchItem.Content"
args22(7).Value = false
args22(8).Name = "SearchItem.AsianOptions"
args22(8).Value = false
args22(9).Name = "SearchItem.AlgorithmType"
args22(9).Value = 1
args22(10).Name = "SearchItem.SearchFlags"
args22(10).Value = 65536
args22(11).Name = "SearchItem.SearchString"
args22(11).Value = " "+CHR$(34)
args22(12).Name = "SearchItem.ReplaceString"
args22(12).Value = " «"
args22(13).Name = "SearchItem.Locale"
args22(13).Value = 255
args22(14).Name = "SearchItem.ChangedChars"
args22(14).Value = 2
args22(15).Name = "SearchItem.DeletedChars"
args22(15).Value = 2
args22(16).Name = "SearchItem.InsertedChars"
args22(16).Value = 2
args22(17).Name = "SearchItem.TransliterateFlags"
args22(17).Value = 1280
args22(18).Name = "SearchItem.Command"
args22(18).Value = 3
args22(19).Name = "SearchItem.SearchFormatted"
args22(19).Value = false
args22(20).Name = "SearchItem.AlgorithmType2"
args22(20).Value = 2
args22(21).Name = "Quiet"
args22(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args22())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args44(21) as new com.sun.star.beans.PropertyValue
args44(0).Name = "SearchItem.StyleFamily"
args44(0).Value = 2
args44(1).Name = "SearchItem.CellType"
args44(1).Value = 0
args44(2).Name = "SearchItem.RowDirection"
args44(2).Value = true
args44(3).Name = "SearchItem.AllTables"
args44(3).Value = false
args44(4).Name = "SearchItem.SearchFiltered"
args44(4).Value = false
args44(5).Name = "SearchItem.Backward"
args44(5).Value = false
args44(6).Name = "SearchItem.Pattern"
args44(6).Value = false
args44(7).Name = "SearchItem.Content"
args44(7).Value = false
args44(8).Name = "SearchItem.AsianOptions"
args44(8).Value = false
args44(9).Name = "SearchItem.AlgorithmType"
args44(9).Value = 1
args44(10).Name = "SearchItem.SearchFlags"
args44(10).Value = 65536
args44(11).Name = "SearchItem.SearchString"
args44(11).Value = CHR$(34)
args44(12).Name = "SearchItem.ReplaceString"
args44(12).Value = "»"
args44(13).Name = "SearchItem.Locale"
args44(13).Value = 255
args44(14).Name = "SearchItem.ChangedChars"
args44(14).Value = 2
args44(15).Name = "SearchItem.DeletedChars"
args44(15).Value = 2
args44(16).Name = "SearchItem.InsertedChars"
args44(16).Value = 2
args44(17).Name = "SearchItem.TransliterateFlags"
args44(17).Value = 1280
args44(18).Name = "SearchItem.Command"
args44(18).Value = 3
args44(19).Name = "SearchItem.SearchFormatted"
args44(19).Value = false
args44(20).Name = "SearchItem.AlgorithmType2"
args44(20).Value = 2
args44(21).Name = "Quiet"
args44(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args44())

end sub