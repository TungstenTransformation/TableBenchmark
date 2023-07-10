# KT Project auto documentation
*created by [KT markdown](https://github.com/KofaxRPA/KT_markdown#kt_markdown)
project version= 23.0
# Class: project
# Class: document
# Class: paystub
 * field: LineItems:
 * field: TableRowCount:
 * field: TableTotalPriceSum:
 * field: TableRowAlignment:
 * field: TableColumnAlignment:
 * field: TableCells:
 * locator: ATL:TableClsLoc
 * locator: FL_Amounts:RegularExpressions
 * locator: TL:Tables
## Project Level Script
```vb
Option Explicit

' Project Script

```
## Script for class 'paystub'
```vb
'#Language "WWB-COM"
Option Explicit

' Class script: invoice

Private Sub Document_AfterExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   Dim Truth As Boolean, TruthDoc As New CscXDocument, TestDoc As New CscXDocument, TestDocFileName As String
   'If the XDoc contains the XValue "OriginalFileName" then it is an online learning sample that came from KTA - so it contains the truth.
   If pXDoc.XValues.ItemExists("OriginalFileName") Then
      Truth = True
      'This is a new sample that came from KTA . We just need to copy the truth back into the original document
      'But the locators just ran and have incorrect values - we need to ignore them and load the file from the file system.
      TruthDoc.Load(pXDoc.FileName)
      'When you drag an xdoc from samples set to test set it is added to a subfolder. The matching original is in the parent directory
      TestDocFileName=Path_GetFolder(Path_GetFolder(pXDoc.FileName)) & pXDoc.XValues.ItemByName("OriginalFileName").Value
      TestDocFileName=Left(TestDocFileName,InStrRev(TestDocFileName,".")) & "xdc"  'OriginalFileName is probably a tif or pdf
      TestDoc.Load(TestDocFileName)
      XDocument_CopyFields(TruthDoc,TestDoc)
      TestDoc.Save()
      Set pXDoc=TestDoc
   End If
   TableBenchmark_Calculate(pXDoc, "LineItems", "Total Price", DefaultAmountFormatter,Truth)
   If Truth Then pXDoc.Save()
End Sub

Sub TableBenchmark_Calculate(pXDoc As CscXDocument, TableFieldName As String, SumColumnName, SumColumnAmountFormatter As CscAmountFormatter,Truth As Boolean)
   'Calculate all the table meta fields for the benchmark.
   Dim Table As CscXDocTable, SumIsValid As Boolean, Field As CscXDocField, FieldName As String, ErrDescription As String, TruthDoc As New CscXDocument, TruthTable As CscXDocTable
   Set Table=pXDoc.Fields.ItemByName(TableFieldName).Table
   TruthDoc.Load(pXDoc.FileName)
   Set Field=pXDoc.Fields.ItemByName("TableRowCount")
   Field.Text=CStr(Table.Rows.Count)
   Field.Confidence=1.00: Field.ExtractionConfident=True
   If SumColumnName<>"" Then
      Set Field=pXDoc.Fields.ItemByName("Table" & Replace(SumColumnName," ","")&"Sum")
      Field.Text=Format(Table.GetColumnSum(Table.Columns.ItemByName(SumColumnName).IndexInTable,SumIsValid),"0.00")
      If SumIsValid Then SumColumnAmountFormatter.FormatField(Field)
      Field.Confidence=1.00: Field.ExtractionConfident=True
   End If
   If Truth Then
      'This is the Truth document - we just set all values to perfect "1.00"
      For Each FieldName In Split("TableRowAlignment TableColumnAlignment TableCells")
         Set Field=pXDoc.Fields.ItemByName(FieldName)
         Field.Text="1.00"
         Field.Confidence=1.00
         Field.ExtractionConfident=True
      Next
   Else
   If Not TruthDoc.Fields.Exists(TableFieldName) Then Exit Sub ' There are no fields in the truth document. nothing to do
   'Here we compare the extracted table with the truth table from the xdoc in the filesystem
      Set TruthTable=TruthDoc.Fields.ItemByName(TableFieldName).Table
      Set Field=pXDoc.Fields.ItemByName("TableRowAlignment")
      Field.Text=Format(Tables_RowAlignment(pXDoc,Table,TruthTable,ErrDescription),"0.00")
      If ErrDescription <> "" Then Field.Text= Field.Text & vbCrLf & " Bad Rows:" & ErrDescription ' so we can see the misaligned row numbers in the benchmark
      Field.Confidence=1.00: Field.ExtractionConfident=True
      ErrDescription=""
      Set Field=pXDoc.Fields.ItemByName("TableColumnAlignment")
      Field.Text=Format(Tables_ColumnAlignment(pXDoc,Table,TruthTable, ErrDescription),"0.00")
      If ErrDescription <> "" Then Field.Text= Field.Text & vbCrLf & " Bad Columns:" & ErrDescription ' so we can see the misaligned column numbers in the benchmark
      Field.Confidence=1.00: Field.ExtractionConfident=True
      Set Field=pXDoc.Fields.ItemByName("TableCells")
      Field.Text=Format(Tables_CompareCells(Table,TruthTable,ErrDescription),"0.00")
      If ErrDescription <> "" Then Field.Text= Field.Text & vbCrLf&  ErrDescription ' so we can see the wrong text in the benchmark 'only show 10 results!
      Field.Confidence=1.00: Field.ExtractionConfident=True
   End If
End Sub

Function Field_Set(pXDoc As CscXDocument, FieldName As String, FieldText As String, Confidence As Double, ErrDescription As String) As CscXDocField
   Dim Field As CscXDocField
   Set Field=pXDoc.Fields.ItemByName(FieldName)
   Field.Text=FieldText
   Field.Confidence=Confidence
   Field.ExtractionConfident=True
   Field.ErrorDescription=ErrDescription
   Return Field
End Function


Function Tables_RowAlignment(pXDoc As CscXDocument, Table As CscXDocTable, TruthTable As CscXDocTable, ByRef ErrDescription As String) As Double
   Dim Alignment As Double, R As Long, TotalAlignment As Double
   ErrDescription=""
   If Table.Rows.Count=0 Then Return 0
   If TruthTable.Rows.Count=0 Then Return 0
   For R=0 To Table.Rows.Count-1
      If R<TruthTable.Rows.Count Then
         Alignment =Rows_Alignment(Table.Rows(R),TruthTable.Rows(R))
         If Alignment <1.00 Then ErrDescription=ErrDescription & CStr(R+1) & ","
         TotalAlignment=TotalAlignment+ Alignment
      End If
   Next
   If ErrDescription<>"" Then ErrDescription= Left(ErrDescription,Len(ErrDescription)-1) 'remove trailing space
   Return TotalAlignment/Max(Table.Rows.Count,TruthTable.Rows.Count) ' returns 1.00 if perfect alignment
End Function

Function Rows_Alignment(Row1 As CscXDocTableRow, Row2 As CscXDocTableRow) As Double
   Dim A As Double, B As Double, Overlap As Double, P As Long, Pages As Long
   'Some rows can page wrap onto another page. It's actually possible for a single row to cover many pages, but unlikely.
   If Row1.StartPage<>Row2.StartPage Then Return 0
   If Row1.EndPage<>Row2.EndPage Then Return 0
   For P=Row1.StartPage To Row1.EndPage
      If Row1.Height(P)>0 And Row2.Height(P)>0 Then
         A=Max(Row1.Top(P)+Row1.Height(P)-Row2.Top(P),0) ' distance from top of row2 to bottom of row1
         B=Max(Row2.Top(P)+Row2.Height(P)-Row1.Top(P),0) ' distance from top of row1 to bottom of row2
         Overlap =Overlap+ Min(A,B)/Max(A,B) ' divide the inside overlap by the outer span. If they are the same, then it gives 1.00
      End If
   Next
   Pages = Max(Row1.EndPage-Row1.StartPage+1,Row2.EndPage-Row2.StartPage+1) ' calculate if any row wraps across one or more pages
   Return Overlap/Pages
End Function

Function Tables_ColumnAlignment(pXDoc As CscXDocument, Table As CscXDocTable,TruthTable As CscXDocTable,ByRef ErrDescription As String) As Double
   Dim Alignment As Double, C As Long
   Dim TotalAlignment As Double
   ErrDescription=""
   If Table.Columns.Count<> TruthTable.Columns.Count Then Return 0 ' these tables are not using the same table model!!!
   For C=0 To Table.Columns.Count-1
      If C<TruthTable.Columns.Count Then
         Alignment=Columns_Alignment(Table.Columns(C),TruthTable.Columns(C),Table)
         If Alignment <1.00 Then ErrDescription=ErrDescription & CStr(C+1) & ","
         TotalAlignment=TotalAlignment+ Alignment
      End If
   Next
   If ErrDescription<>"" Then ErrDescription= Left(ErrDescription,Len(ErrDescription)-1) 'remove trailing space
   Return TotalAlignment/Table.Columns.Count ' returns 1.00 if perfect alignment
End Function


Function Columns_Alignment(Column1 As CscXDocTableColumn, Column2 As CscXDocTableColumn, Table As CscXDocTable) As Double
   Dim A As Double, B As Double, Overlap As Double, P As Long, Pages As Long, StartPage As Long, EndPage As Long
   If Column1.StartPage<>Column2.StartPage Then Return 0
   If Column1.EndPage<>Column2.EndPage Then Return 0
   StartPage=Table.Rows(0).StartPage 'There is a bug that Column.StartPage and Column.EndPage are always -1, so i need to read from rows.
   EndPage=Table.Rows(Table.Rows.Count-1).EndPage
   For P= StartPage To EndPage
      If Column1.Width(P)=0 And Column2.Width(P)=0 Then
         Overlap=Overlap+1' we allow empty columns
      Else
         A=Max(Column1.Left(P)+Column1.Width(P)-Column2.Left(P),0)
         B=Max(Column2.Left(P)+Column2.Width(P)-Column1.Left(P),0)
         Overlap=Overlap+Min(A,B)/Max(A,B)
      End If
   Next
   Pages = Max(Column1.EndPage-Column1.StartPage+1,Column2.EndPage-Column2.StartPage+1) ' calculate how many pages
   Return Overlap/Pages
End Function

Function Tables_CompareCells(Table As CscXDocTable, TruthTable As CscXDocTable, ByRef ErrDescription As String) As Double
   Dim R As Long, C As Long, Cell As CscXDocTableCell, TruthCell As CscXDocTableCell, Errors As Long
   Const MAXERRORS=10 'only show this many errors
   'Check that all the table cells agree
   ErrDescription=""
   If Table.Columns.Count<>TruthTable.Columns.Count Then
      ErrDescription = "Tables should have same table models"
      Return 0
   End If
   For R=0 To Table.Rows.Count-1
      For C=0 To Table.Columns.Count-1
         If R<TruthTable.Rows.Count Then
            Set Cell=Table.Rows(R).Cells(C)
            Set TruthCell=TruthTable.Rows(R).Cells(C)
            If Cell.Text<>TruthCell.Text Then
               If Errors <MAXERRORS Then
                  ErrDescription= ErrDescription & vbCrLf & "R" & CStr(R+1) & "C" & CStr(C+1) & ":  " & String_Truncate(Cell.Text) & vbCrLf & Space(12) &"[" & String_Truncate(TruthCell.Text) & "]"
               End If
               Errors = Errors +1
            End If
         End If
      Next
   Next
   ErrDescription = "Total Cell Errors: " & CStr(Errors) & vbCrLf & ErrDescription
   Return 1-Errors/Table.Rows.Count/Table.Columns.Count
End Function

Function String_Truncate(A As String) As String
   Const MAXTEXT=35 'truncate all text to this many characters
   Return Left(A,MAXTEXT) & IIf(Len(A)>MAXTEXT, ".","")
End Function


Function Min(A,B) 'typeless function works with all variable types
   Return IIf(A<B,A,B)
End Function

Function Max(A,B)
   Return IIf(A>B,A,B)
End Function

Sub XDocument_CopyFields(A As CscXDocument, B As CscXDocument)
   Dim F As Long, XScale As Double, YScale As Double
   'KTA may have changed the resolution of the document, so we need to scale the pixels. (e.g. original is 300 dpi, and KTA's Scan profile made everything 200 dpi.)
   XScale=B.CDoc.Pages(0).XRes/A.CDoc.Pages(0).XRes
   YScale=B.CDoc.Pages(0).YRes/A.CDoc.Pages(0).YRes
   For F=0 To A.Fields.Count-1
      If B.Fields.Exists(A.Fields(F).Name) Then Field_Copy(A.Fields(F),B.Fields(F),XScale,YScale)
   Next
End Sub

Function Path_GetFolder(PathName As String) As String
   'Return the ParentFolder
   If Right(PathName,1)="\" Then PathName=Left(PathName,Len(PathName)-1)
   Return Left(PathName,InStrRev(PathName,"\"))
End Function

Sub Field_Copy(A As CscXDocField,B As CscXDocField,XScale As Double, YScale As Double) 'copy a field or a table
   Dim R As Long, ARows As CscXDocTableRows, BRows As CscXDocTableRows, C As Long
   Select Case A.FieldType
   Case CscExtractionFieldType.CscFieldTypeSimpleField
      B.PageIndex=A.PageIndex
      B.Left=A.Left*XScale
      B.Top=A.Top*YScale
      B.Width=A.Width*XScale
      B.Height=A.Height*YScale
      B.Confidence=A.Confidence
      B.ExtractionConfident=A.ExtractionConfident
      B.Valid=A.Valid
      B.DoubleValue=A.DoubleValue
      B.DateValue=A.DateValue
      B.DateFormatted=A.DateFormatted
      B.DoubleFormatted=A.DoubleFormatted
   Case CscExtractionFieldType.CscFieldTypeTable
      Set ARows=A.Table.Rows
      Set BRows=B.Table.Rows
      BRows.Clear
      For R=0 To ARows.Count-1
         BRows.Append
         For C = 0 To A.Table.Columns.Count-1
            TableCell_Copy(ARows(R).Cells(C),BRows(R).Cells(C),XScale, YScale)
         Next
      Next
   Case Else
   End Select
End Sub

Sub TableCell_Copy(A As CscXDocTableCell, B As CscXDocTableCell,XScale As Double, YScale As Double) 'copy a single table cell
   Dim Word As New CscXDocWord
   Word.PageIndex=A.PageIndex
   Word.Left=A.Left*XScale
   Word.Top=A.Top*YScale
   Word.Width=A.Width*XScale
   Word.Height=A.Height*YScale
   Word.Text=A.Text
   B.AddWordData(Word)
   B.ExtractionConfident=True
   B.Valid=True
   Set Word=Nothing
End Sub



```
## Formatters  
DateFormatter : DateFormatter  
AmountFormatter : AmountFormatter  
*Default Date   Formatter*: DefaultDateFormatter  
*Default Amount Formatter*: DefaultAmountFormatter  
## Databases  
## Dictionaries  
## Table Settings  
Global Column 0 : Position  
Global Column 1 : Quantity  
Global Column 2 : Description  
Global Column 3 : Unit Price  
Global Column 4 : Total Price  
Global Column 5 : Discount  
Global Column 6 : Unit Measure  
Global Column 7 : Article Code  
Global Column 8 : Supplier Article Code  
Global Column 9 : Order Number  
Global Column 10 : Delivery Note Number  
Global Column 11 : Tax Rate  
Global Column 10000 : PO Item Number  
Global Column 15000 : Tax Amount  
Global Column 15001 : Discount Amount  
Global Column 16000 : Match Remark  
Global Column 20001 : YearToDate  
Table Model: Earnings  