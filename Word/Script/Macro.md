Sure, here are 10 VBA macro scripts for Word that demonstrate various functionalities:

1. **Apply Specific Formatting to Selected Text**:
   ```vba
   Sub ApplyFormatting()
       With Selection.Font
           .Name = "Arial"
           .Size = 12
           .Bold = True
           .Color = RGB(0, 0, 255) ' Blue color
       End With
   End Sub
   ```

2. **Insert Current Date and Time at Cursor Position**:
   ```vba
   Sub InsertDateTime()
       Selection.InsertDateTime DateTimeFormat:="dd/MM/yyyy HH:mm:ss", InsertAsField:=False
   End Sub
   ```

3. **Create and Insert Table with Specified Number of Rows and Columns**:
   ```vba
   Sub InsertTable()
       Dim tbl As Table
       Set tbl = Selection.Tables.Add(Selection.Range, NumRows:=3, NumColumns:=3)
   End Sub
   ```

4. **Find and Replace Text within the Document**:
   ```vba
   Sub FindAndReplace()
       Selection.Find.ClearFormatting
       Selection.Find.Execute FindText:="oldText", ReplaceWith:="newText", Replace:=wdReplaceAll
   End Sub
   ```

5. **Insert Bookmark at Current Selection**:
   ```vba
   Sub InsertBookmark()
       Selection.Bookmarks.Add Name:="MyBookmark", Range:=Selection.Range
   End Sub
   ```

6. **Count Words in Document and Display in Message Box**:
   ```vba
   Sub CountWords()
       Dim wordCount As Integer
       wordCount = ActiveDocument.Words.Count
       MsgBox "Total words: " & wordCount
   End Sub
   ```

7. **Print Document with Specified Printer and Settings**:
   ```vba
   Sub PrintWithSettings()
       Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
           wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
           Collate:=True, Background:=True, PrintToFile:=False, PrintZoomColumn:=0, _
           PrintZoomRow:=0, PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0
   End Sub
   ```

8. **Insert Hyperlink at Cursor Position**:
   ```vba
   Sub InsertHyperlink()
       ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="https://www.example.com", _
           SubAddress:="", ScreenTip:="Visit Example Website", TextToDisplay:="Example Website"
   End Sub
   ```

9. **Open a Dialog Box to Save Document**:
   ```vba
   Sub SaveAsDialog()
       Dialogs(wdDialogFileSaveAs).Show
   End Sub
   ```

10. **Highlight Text with a Specific Color**:
    ```vba
    Sub HighlightText()
        Selection.Range.HighlightColorIndex = wdYellow
    End Sub
    ```

These VBA macros can be added to Word's Visual Basic Editor and executed to perform their respective tasks within a Word document.
