
Attribute VB_Name = "DeliveryMergeModule"

Sub RunDeliveryMerge(savePath As String, currentDate As Date, _
                     excelPath As String, wordTemplatePath As String, _
                     keyword As String, filePrefix As String)

    Dim ExlApp As Object, ExlWb As Object, ExlWs As Object, SearchRange As Object
    Dim WordApp As Object, WordDoc As Object, cell As Object
    Dim Filename As String

    Set ExlApp = CreateObject("Excel.Application")
    ExlApp.Visible = False
    Set ExlWb = ExlApp.Workbooks.Open(excelPath)
    Set ExlWs = ExlWb.Sheets(1)
    Set SearchRange = ExlWs.Range("B2:B999")

    For Each cell In SearchRange
        If InStr(cell.Value, keyword) > 0 Then
            Set WordApp = CreateObject("Word.Application")
            WordApp.DisplayAlerts = False
            Set WordDoc = WordApp.Documents.Open(wordTemplatePath)
            WordApp.Visible = True

            ExlWb.Close SaveChanges:=False
            ExlApp.Quit
            Exit For
        End If
    Next

    If Not WordDoc Is Nothing Then
        With WordDoc.MailMerge
            .Destination = wdSendToNewDocument
            .Execute
        End With

        Filename = filePrefix & Format(currentDate, "MMdd")

        WordApp.ActiveDocument.ExportAsFixedFormat OutputFileName:=savePath & Filename & ".pdf", ExportFormat:=17
        WordApp.ActiveDocument.SaveAs savePath & Filename & ".doc", 0

        WordDoc.Close SaveChanges:=False
        WordApp.Quit
    Else
        ExlWb.Close SaveChanges:=False
        ExlApp.Quit
    End If

    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set ExlApp = Nothing
    Set ExlWb = Nothing
    Set ExlWs = Nothing
End Sub

Sub Main_DeliveryMerge()
    Dim currentDate As Date
    currentDate = Date - 1

    Dim basePath As String
    basePath = ThisWorkbook.Path & "\"

    Call RunDeliveryMerge(basePath & "output\", currentDate, _
        basePath & "delivery_data_example.xlsx", _
        basePath & "template_unit01.docx", _
        "Unit01", "Acceptance_Unit01_")

    Call RunDeliveryMerge(basePath & "output\", currentDate, _
        basePath & "delivery_data_example.xlsx", _
        basePath & "template_unit02.docx", _
        "Unit02", "Acceptance_Unit02_")

    MsgBox "All merges completed!"
End Sub
