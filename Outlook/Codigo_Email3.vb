Sub MandaEmail()
    
    Dim EnviarPara As String
    Dim Assunto As String
    Dim Cel As Range, Fim As Long, i As Byte, Hora() As Date, Tempo As Date
    Dim Mensagem As String, Hoje As Date, Caminho As String, Caminho2 As String, Caminho3 As String, Assinatura As String
    
            'Atribuir valores
            Planilha1.Select
            Fim = Range("A" & Rows.Count).End(xlUp).Row
            Tempo = Format(Now, "hh:mm:ss")
            Hoje = Format(Now, "dd/mm/yyyy")
            i = 0
            
            'Loop para contar a qtd programada do dia
            For Each Cel In Range("A2:A" & Fim)
                If Cel.Value = Hoje Then
                i = i + 1
                End If
            Next Cel
            
            'Definindo o tamanho dos Array
            ReDim Hora(i)
            
            i = 0
            
            'Loop para armazenar as horas
            For Each Cel In Range("A2:A" & Fim)
                If Cel.Value = Hoje Then
                    Cel.Activate
                    Hora(i) = Cel.Offset(0, 1).Value
                    Hora(i) = Format(Hora(i), "hh:mm:ss")
                    i = i + 1
                End If
            Next Cel
            
            i = 0
            
            'Loop para a leitura da planilha
            For Each Cel In Range("A2:A" & Fim)
                If Cel.Value = Hoje Then
                If Hora(i) = Tempo Then
                    Cel.Activate
                    Assunto = Cel.Offset(0, 2).Value
                    EnviarPara = Cel.Offset(0, 3).Value
                    Mensagem = Cel.Offset(0, 4).Value
                    Assinatura = Cel.Offset(0, 5).Value
                    Caminho = Cel.Offset(0, 6).Value
                    Caminho2 = Cel.Offset(0, 7).Value
                    Caminho3 = Cel.Offset(0, 8).Value
                    Application.DisplayAlerts = False
                    'Guardar informações na aba de histórico
                    Planilha1.Select
                    Range(ActiveCell.Offset(0, 0), Selection.End(xlToRight)).Copy
                    Planilha3.Select
                    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
                    ActiveCell.PasteSpecial
                    Planilha3.Range("A" & Rows.Count).End(xlDown).End(xlUp).Select
                    Planilha1.Select
                    Range(ActiveCell.Offset(0, 0), Selection.End(xlToRight)).Delete
                    Application.DisplayAlerts = True
                    End If
                    End If
                    i = i + 1
            Next Cel
            Range("A2").Select
            
            'Chamando o metodo com os parâmetros
            Envia_Emails EnviarPara, Assunto, Mensagem, Caminho, Caminho2, Caminho3, Assinatura
End Sub

Sub Envia_Emails(EnviarPara As String, Assunto As String, Mensagem As String, Caminho As String, Caminho2 As String, Caminho3 As String, Assinatura As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim strArray() As String, strArrayAss() As String
    Dim intCount As Integer
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    On Error GoTo Desvio
    
    With OutlookMail
        .To = EnviarPara
        .CC = ""
        .BCC = ""
        .Importance = olImportanceHigh
        .Subject = Assunto
    'Array para armazenar os dados separar pelo caractere
    strArray = Split(Mensagem, "<>")
    strArrayAss = Split(Assinatura, "<>")
    'Limpa String
        Mensagem = Empty
        Assinatura = Empty
        
    'Loop para quebrar linha da mensagem no outlook
    For intCount = LBound(strArray) To UBound(strArray)
        If Mensagem = Empty Then
            Mensagem = strArray(intCount)
        End If
        If intCount < 1 Then
            Mensagem = Mensagem & "<BR>"
            Else
            Mensagem = Mensagem & "<BR>" & strArray(intCount)
        End If
    Next
    
    'Incluir assinatura
    For intCount = LBound(strArrayAss) To UBound(strArrayAss)
        Assinatura = Assinatura & "<BR>" & strArrayAss(intCount)
    Next
    
        'Ação para colocar a mensagem, assinatura e a imagem no corpo de email
        .HTMLBody = Mensagem & "<BR><BR>" & Assinatura & "<BR>" & _
       "<img src='C:\Users\Inmetrics\Desktop\Assinatura.png'>"
    
       'Condição para enviar o anexo
       If Caminho = Empty Then
            Caminho = "Sem anexo"
            Else
            .Attachments.Add Caminho
       End If
        
       If Caminho2 = Empty Then
            Caminho2 = "Sem anexo"
            Else
            .Attachments.Add Caminho2
       End If
       
       If Caminho3 = Empty Then
            Caminho3 = "Sem anexo"
            Else
            .Attachments.Add Caminho3
       End If
        
        .Send
        '.Display
    End With
    
    'Mensagem de confirmação
    MsgBox "Email Enviado com Sucesso!" & vbCr & vbCr & _
    "Para: " & EnviarPara & vbCr & _
    "Às: " & Format(Time, "hh:mm:ss") & vbCr & _
    "Assunto: " & Assunto, vbInformation, "Envio de Emails - Erick Automação"
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    Exit Sub
    
'Trata erro
Desvio:
    MsgBox Err.Number & " - " & Err.Description

    
End Sub

Sub Executar()

Dim Cel As Range, Fim As Long, Hoje As String, Cont As Byte, i As Byte, Max As Byte
Dim Hora() As Date, Assunto As String, Tempo As Date, Caminho As String

Planilha1.Select
Fim = Range("A" & Rows.Count).End(xlUp).Row
Hoje = Format(Now, "dd/mm/yyyy")

Application.StatusBar = "Aguarde a execução do programa..."
Application.ScreenUpdating = False

      
'Validações necessárias para execução
For Each Cel In Range("A2:A" & Fim)

        If Cel.Value = Empty Then
            MsgBox "Informe uma data!", vbCritical
            Exit Sub
        End If
            
        If Cel.Offset(0, 1).Value = Empty Then
            MsgBox "Informe um horario!", vbCritical
            Exit Sub
        End If
        
        If Cel.Offset(0, 2).Value = Empty Then
            MsgBox "Informe o assunto!", vbCritical
            Exit Sub
        End If
        
        If Cel.Offset(0, 3).Value = Empty Then
            MsgBox "Informe quem irá receber o email!", vbCritical
            Exit Sub
        End If
    
        If Cel.Offset(0, 4).Value = Empty Then '
            MsgBox "Informe a mensagem!", vbCritical
            Exit Sub
        End If
        
        If Cel.Value = Hoje Then
            Cont = Cont + 1
        End If
Next Cel

'Definir tamanho da Array
ReDim Hora(Cont - 1)

i = 0

'Loop para armazenar as horas
For Each Cel In Range("A2:A" & Fim)

         If Cel.Value = Hoje Then
            Cel.Activate
            Hora(i) = Cel.Offset(0, 1).Value
            Hora(i) = Format(Hora(i), "hh:mm:ss")
            i = i + 1
         End If
Next Cel


Range("A2").Select
Max = i - 1

'Loop para programar a aplicação
For i = 0 To Max

    Tempo = Format(Hoje & Space(1) & Hora(i), "dd/mm/yyyy hh:mm:ss")
    MsgBox "Email Programado para: " & vbCr & vbCr & _
    Tempo, vbInformation, "Programação de Emails - Erick Automação"

'MsgBox Format(Tempo, "dd/mm/yyyy hh:mm:ss")

    'If MsgBox("Você deseja anexar algum arquivo?", vbQuestion + vbYesNo + vbDefaultButton2, "Diretorio do arquivo") = vbYes Then
            'Caminho = InputBox("Digite o caminho do anexo: ", "Diretorio do arquivo", "caminho")
            'Else
               'MsgBox "Email sem anexo!", vbInformation, "Importante"
               'Caminho = Empty
            'End If

    Application.OnTime TimeValue(Tempo), "MandaEmail"
Next i

Application.StatusBar = Empty
Application.ScreenUpdating = True

End Sub
