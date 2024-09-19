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
    
    ' K�saltma modelini tan�mla (B�y�k harflerden olu�an 2 ve �zeri karakterli kelimeler)
    Set kisaltmalar = CreateObject("Scripting.Dictionary")
    Set kisaltmaList = CreateObject("Scripting.Dictionary")
    paragraphText = doc.Content.text
    
    ' K�saltmalar� bul (Regex kullan�m� i�in vbScript Regex mod�l�n� aktif et)
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
    excelPath = "dosyaYolu\kisaltmalar.xlsx" ' Burada dosyan�n tam yolunu verin.
    
    ' Excel dosyas�n� a�
    Set excelApp = CreateObject("Excel.Application")
    Set excelWorkbook = excelApp.Workbooks.Open(excelPath)
    Set excelSheet = excelWorkbook.Sheets(1)
    
    ' Excel'den k�saltmalar� ve a��klamalar� al
    Set kisaltmaDict = CreateObject("Scripting.Dictionary")
    i = 2 ' 1. sat�r ba�l�klar i�in ayr�lm��, 2. sat�rdan itibaren veriler var
    Do While excelSheet.Cells(i, 1).Value <> ""
        If kisaltmaDict.Exists(excelSheet.Cells(i, 1).Value) Then
            ' E�er k�saltma zaten varsa, a��klamay� mevcut listeye ekle
            kisaltmaDict(excelSheet.Cells(i, 1).Value) = kisaltmaDict(excelSheet.Cells(i, 1).Value) & "; " & excelSheet.Cells(i, 2).Value
        Else
            ' Yeni bir k�saltma ekle
            kisaltmaDict.Add excelSheet.Cells(i, 1).Value, excelSheet.Cells(i, 2).Value
        End If
        i = i + 1
    Loop
    
    excelWorkbook.Close False
    excelApp.Quit
    Set excelApp = Nothing
    
    ' E�le�en k�saltmalar� bul ve bilgileri topla
    Set kisaltmaMatch = CreateObject("Scripting.Dictionary")
    For Each kisaltma In kisaltmalar
        If kisaltmaDict.Exists(kisaltma) Then
            ' E�er ayn� k�saltmaya birden fazla a��klama varsa
            If InStr(kisaltmaDict(kisaltma), ";") > 0 Then
                explanationOptions = Split(kisaltmaDict(kisaltma), ";")
                selectedExplanation = ""
                
                ' Se�enekleri kullan�c�ya sun ve se�im yapmas�n� iste
                optionsText = ""
                For j = LBound(explanationOptions) To UBound(explanationOptions)
                    optionsText = optionsText & (j + 1) & ": " & Trim(explanationOptions(j)) & vbCrLf
                Next j
                
                Do
                    userChoice = InputBox("A�a��daki a��klamalardan birini se�in:" & vbCrLf & optionsText, "K�saltma: " & kisaltma)
                    
                    ' Kullan�c� se�imini kontrol et
                    If IsNumeric(userChoice) And CInt(userChoice) > 0 And CInt(userChoice) <= UBound(explanationOptions) + 1 Then
                        selectedExplanation = Trim(explanationOptions(CInt(userChoice) - 1))
                    Else
                        MsgBox "Ge�ersiz se�im. L�tfen ge�erli bir say� girin."
                    End If
                Loop Until selectedExplanation <> ""
                
                ' Se�ilen a��klamay� s�zl��e ekle
                kisaltmaMatch.Add kisaltma, selectedExplanation
            Else
                ' Tek a��klama varsa do�rudan ekle
                kisaltmaMatch.Add kisaltma, kisaltmaDict(kisaltma)
            End If
        End If
    Next kisaltma
    
    ' Metnin sonuna git ve yeni sayfa ekle
    Selection.EndKey Unit:=wdStory ' �mleci metnin sonuna ta��
    Selection.InsertBreak Type:=wdPageBreak ' Yeni sayfa ekle
    
    ' Ba�l��� ekle ve ortala
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Selection.TypeText "KISALTMALAR VE A�IKLAMALARI" & vbCrLf & vbCrLf
    
    ' Bo�luk b�rak (iki sat�r bo�luk)
    Selection.TypeParagraph
    Selection.TypeParagraph
    
    ' K�saltmalar� ve a��klamalar� ekle, yaz� tipi k���k ve sola yasl� olacak
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Bold = False
    Selection.Font.Size = 10
    
    For Each kisaltma In kisaltmaMatch
        Selection.TypeText kisaltma & ": " & kisaltmaMatch(kisaltma) & vbCrLf
    Next kisaltma
End Sub

