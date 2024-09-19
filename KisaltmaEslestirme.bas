Attribute VB_Name = "NewMacros"
Sub KisaltmaEslestirme()
    Dim doc As Document
    Dim kisaltmalar As Object
    Dim kisaltmaList As Object
    Dim kisaltma As Variant
    Dim i As Long
    Dim excelPath As String
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim kisaltmaDict As Object
    Dim kisaltmaMatch As Object
    Dim paragraphText As String
    Dim userChoice As String
    Dim selectedExplanation As String
    Dim explanationOptions() As String
    Dim optionsText As String
    Dim j As Integer
    
    ' Aktif Word belgesi
    Set doc = ActiveDocument
    
    ' Kýsaltma modelini tanýmla (Büyük harflerden oluþan 2 ve üzeri karakterli kelimeler)
    Set kisaltmalar = CreateObject("Scripting.Dictionary")
    Set kisaltmaList = CreateObject("Scripting.Dictionary")
    paragraphText = doc.Content.text
    
    ' Kýsaltmalarý bul (Regex kullanýmý için vbScript Regex modülünü aktif et)
    With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = "\b[A-Z]{2,}\b"
        If .test(paragraphText) Then
            For Each kisaltma In .Execute(paragraphText)
                If Not kisaltmalar.Exists(kisaltma.Value) Then
                    kisaltmalar.Add kisaltma.Value, kisaltma.Value
                End If
            Next kisaltma
        End If
    End With
    
    ' Sabit Excel dosya yolunu belirtin
    excelPath = "dosyaYolu\kisaltmalar.xlsx" ' Burada dosyanýn tam yolunu verin.
    
    ' Excel dosyasýný aç
    Set excelApp = CreateObject("Excel.Application")
    Set excelWorkbook = excelApp.Workbooks.Open(excelPath)
    Set excelSheet = excelWorkbook.Sheets(1)
    
    ' Excel'den kýsaltmalarý ve açýklamalarý al
    Set kisaltmaDict = CreateObject("Scripting.Dictionary")
    i = 2 ' 1. satýr baþlýklar için ayrýlmýþ, 2. satýrdan itibaren veriler var
    Do While excelSheet.Cells(i, 1).Value <> ""
        If kisaltmaDict.Exists(excelSheet.Cells(i, 1).Value) Then
            ' Eðer kýsaltma zaten varsa, açýklamayý mevcut listeye ekle
            kisaltmaDict(excelSheet.Cells(i, 1).Value) = kisaltmaDict(excelSheet.Cells(i, 1).Value) & "; " & excelSheet.Cells(i, 2).Value
        Else
            ' Yeni bir kýsaltma ekle
            kisaltmaDict.Add excelSheet.Cells(i, 1).Value, excelSheet.Cells(i, 2).Value
        End If
        i = i + 1
    Loop
    
    excelWorkbook.Close False
    excelApp.Quit
    Set excelApp = Nothing
    
    ' Eþleþen kýsaltmalarý bul ve bilgileri topla
    Set kisaltmaMatch = CreateObject("Scripting.Dictionary")
    For Each kisaltma In kisaltmalar
        If kisaltmaDict.Exists(kisaltma) Then
            ' Eðer ayný kýsaltmaya birden fazla açýklama varsa
            If InStr(kisaltmaDict(kisaltma), ";") > 0 Then
                explanationOptions = Split(kisaltmaDict(kisaltma), ";")
                selectedExplanation = ""
                
                ' Seçenekleri kullanýcýya sun ve seçim yapmasýný iste
                optionsText = ""
                For j = LBound(explanationOptions) To UBound(explanationOptions)
                    optionsText = optionsText & (j + 1) & ": " & Trim(explanationOptions(j)) & vbCrLf
                Next j
                
                Do
                    userChoice = InputBox("Aþaðýdaki açýklamalardan birini seçin:" & vbCrLf & optionsText, "Kýsaltma: " & kisaltma)
                    
                    ' Kullanýcý seçimini kontrol et
                    If IsNumeric(userChoice) And CInt(userChoice) > 0 And CInt(userChoice) <= UBound(explanationOptions) + 1 Then
                        selectedExplanation = Trim(explanationOptions(CInt(userChoice) - 1))
                    Else
                        MsgBox "Geçersiz seçim. Lütfen geçerli bir sayý girin."
                    End If
                Loop Until selectedExplanation <> ""
                
                ' Seçilen açýklamayý sözlüðe ekle
                kisaltmaMatch.Add kisaltma, selectedExplanation
            Else
                ' Tek açýklama varsa doðrudan ekle
                kisaltmaMatch.Add kisaltma, kisaltmaDict(kisaltma)
            End If
        End If
    Next kisaltma
    
    ' Metnin sonuna git ve yeni sayfa ekle
    Selection.EndKey Unit:=wdStory ' Ýmleci metnin sonuna taþý
    Selection.InsertBreak Type:=wdPageBreak ' Yeni sayfa ekle
    
    ' Baþlýðý ekle ve ortala
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Selection.TypeText "KISALTMALAR VE AÇIKLAMALARI" & vbCrLf & vbCrLf
    
    ' Boþluk býrak (iki satýr boþluk)
    Selection.TypeParagraph
    Selection.TypeParagraph
    
    ' Kýsaltmalarý ve açýklamalarý ekle, yazý tipi küçük ve sola yaslý olacak
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Bold = False
    Selection.Font.Size = 10
    
    For Each kisaltma In kisaltmaMatch
        Selection.TypeText kisaltma & ": " & kisaltmaMatch(kisaltma) & vbCrLf
    Next kisaltma
End Sub

