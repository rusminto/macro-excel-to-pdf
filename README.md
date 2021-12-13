# macro-excel-to-pdf

Example of implementation macro to:  
1. get data from excel
  ```vba
  Dim wb As Workbook
  Set wb = ActiveWorkbook

  Dim ws As Worksheet
  Set ws = wb.Sheets("Sheet1")
  
  n = ws.Range("A:A").Find(what:="*", searchdirection:=xlPrevious).Row
  
  for j = 2 To n
    nama = ws.Range("A" & j).Value
  Next
  ```
2. insert excel's data to microsoft word'template
  ```vba
  Set wordApp = CreateObject("Word.Application")
  
  Set wordDoc = wordApp.Documents.Open("D:\stechoq\pdfPassword\template2.docx")
  
  With wordDoc.Content.Find
      .Execute FindText:="<<nama>>", ReplaceWith:=nama, Replace:=wdReplaceAll
  End With
  ```
3. export word to pdf
   ```vba
   rawPdf = "D:\stechoq\pdfPassword\" & ws.Range("A" & j).Value & "_raw.pdf"
   passwordedPdf = "D:\stechoq\pdfPassword\" & ws.Range("A" & j).Value & ".pdf"
        
   wordDoc.ExportAsFixedFormat rawPdf, wdExportFormatPDF
        
   wordDoc.Close (wdDoNotSaveChanges)
   ```
4. add password on selected pdf
   ```vba
   Dim wsh As Object
   Set wsh = VBA.CreateObject("WScript.Shell")
   Dim waitOnReturn As Boolean: waitOnReturn = True
   Dim windowStyle As Integer: windowStyle = 1
   
   fTemp = """" & rawPdf & """"
   oPdf = """" & passwordedPdf & """"
   pwd = """" & password & """"

   cmdStr = "pdftk " & fTemp & " Output " & oPdf & " User_pw " & pwd & " Allow AllFeatures"

   wsh.Run "cmd.exe /S /C " & cmdStr, windowStyle, waitOnReturn

   Kill Replace(fTemp, """", "")
   ```
