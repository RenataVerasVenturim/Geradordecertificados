Use this document to GitHub understand this project is VBA´s project

VBA Codes in this project:
<hr>
Private Sub WorkBook_Open()

Application.Visible = True

Dados_certificado.Show

End Sub
<hr>
'Criar espera entre uma macro e outra
Sub esperar()

newhour = Hour(Now())
newminute = Minute(Now())
'determinar espera em segundos
newsecond = Second(Now()) + 5

waittime = TimeSerial(newhour, newminute, newsecond)

Application.Wait waittime

End Sub
<hr>

Sub criar_novo_certificado()

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8
    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(2, colTab).Value
    
Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(2, 1).Value & ".docx")
arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

' segundo
If Cells(3, 1) <> "" Then
Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

    
    
For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(3, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(3, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'terceiro

If Cells(4, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(4, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(4, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If


'quarto

If Cells(5, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(5, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(5, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quinto

If Cells(6, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(6, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(6, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'sexto

If Cells(7, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(7, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(7, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'sétimo

If Cells(8, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(8, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(8, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'oitavo

If Cells(9, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(9, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(9, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'nono

If Cells(10, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(10, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(10, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo

If Cells(11, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(11, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(11, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo primeiro

If Cells(12, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(12, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(12, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo segundo

If Cells(13, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(13, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(13, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo terceiro

If Cells(14, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(14, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(14, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo quarto

If Cells(15, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(15, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(15, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo quinto

If Cells(16, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(16, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(16, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo sexto

If Cells(15, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(15, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(15, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo sétimo

If Cells(18, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(18, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(18, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo oitavo

If Cells(19, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(19, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(19, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'décimo nono

If Cells(20, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(20, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(20, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo

If Cells(21, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(21, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(21, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo primeiro

If Cells(22, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(22, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(22, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo segundo

If Cells(23, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(23, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(23, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo terceiro

If Cells(24, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(24, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(24, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo quarto

If Cells(25, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(25, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(25, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo quinto

If Cells(25, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(25, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(25, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If


'vigésimo sexto

If Cells(27, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(27, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(27, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo sétimo

If Cells(28, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(28, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(28, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo oitavo

If Cells(29, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(29, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(29, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'vigésimo nono

If Cells(30, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(30, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(30, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo

If Cells(31, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(31, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(31, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo primeiro

If Cells(32, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(32, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(32, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo segundo

If Cells(33, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(33, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(33, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo quarto

If Cells(35, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(35, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(35, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo quinto

If Cells(36, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(36, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(36, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo sexto

If Cells(37, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(37, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(37, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo sétimo

If Cells(38, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(38, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(38, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo oitavo

If Cells(39, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(39, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(39, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'trigésimo nono

If Cells(40, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(40, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(40, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo

If Cells(41, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(41, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(41, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo primeiro

If Cells(42, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(42, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(42, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo segundo

If Cells(43, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(43, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(43, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo terceiro

If Cells(44, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(44, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(44, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo quarto

If Cells(45, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(45, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(45, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo sexto

If Cells(47, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(47, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(47, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo sétimo

If Cells(48, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(48, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(48, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo oitavo

If Cells(49, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(49, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(49, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quadragésimo nono

If Cells(50, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(50, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(50, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

'quinquagésimo

If Cells(51, 1) <> "" Then

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo de certificado" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 8

    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(51, colTab).Value
  

Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "\Certificados" & "\Certificado - " & Cells(51, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

Call esperar

End If

MsgBox ("Certificados gerados e salvos com sucesso na pasta'Certificados'!")

End Sub
<hr>
Sub abrir_formulario()

Dados_certificado.Show

End Sub
<hr>
Sub PDF()
'
'CEntralzar células
Columns("A:H").Select
    Range("H1").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("I7").Select
    
' deixar em paisagem
ActiveSheet.PageSetup.Orientation = xlLandscape

' PDF MACRO

pasta = ThisWorkbook.Path & "\Relatório" & "\Relatório" & ".pdf"
linha = Cells(1048576, 1).End(xlUp).Row + 1
    Range(Cells(1, 1), Cells(linha, 8)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        pasta, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
        
        
End Sub



