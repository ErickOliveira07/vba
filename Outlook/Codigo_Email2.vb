Sub MandaEmail()
    
    Dim EnviarPara As String
    Dim Assunto As String
    Dim Cel As Range, Fim As Long, i As Byte
    Dim Hora As String, Mensagem As String, Hoje As Date, Caminho As String, Assinatura As String
    
            Planilha1.Select
            Fim = Range("A" & Rows.Count).End(xlUp).Row
            Hoje = Format(Now, "dd/mm/yyyy")
            i = 2
            
            'Loop para a leitura da planilha
            For Each Cel In Range("A2:A" & Fim)
                If Cel.Value = Hoje Then
                    Cel.Activate
                    Assunto = Cel.Offset(0, 2).Value
                    EnviarPara = Cel.Offset(0, 3).Value
                    Mensagem = Cel.Offset(0, 4).Value
                    Assinatura = Cel.Offset(0, 5).Value
                    Caminho = Cel.Offset(0, 6).Value
                    'Range(ActiveCell.Offset(0, 0), Selection.End(xlToRight)).Delete
                    End If
                    i = i + 1
            Next Cel
            Range("A2").Select
            
            Envia_Emails EnviarPara, Assunto, Mensagem, Caminho, Assinatura
End Sub

Sub Envia_Emails(EnviarPara As String, Assunto As String, Mensagem As String, Caminho As String, Assinatura As String)
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
        
        .Send
        '.Display
    End With
    
'Trata erro
Desvio:
    MsgBox Err.Number & " - " & Err.Description
    
    Exit Sub
    
    'Mensagem de confirmação
    MsgBox "Email Enviado com Sucesso!" & vbCr & vbCr & _
    "Para: " & EnviarPara & vbCr & _
    "Às: " & Format(Time, "hh:mm:ss") & vbCr & _
    "Assunto: " & Assunto, vbInformation, "Envio de Emails"
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

Sub Executar()

Dim Cel As Range, Fim As Long, Hoje As String, i As Byte
Dim Hora As Date, Assunto As String, Tempo As Date, Caminho As String

Planilha1.Select
Fim = Range("A" & Rows.Count).End(xlUp).Row
Hoje = Format(Now, "dd/mm/yyyy")
i = 2

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
    
        If Cel.Offset(0, 4).Value = Empty Then
            MsgBox "Informe a mensagem!", vbCritical
            Exit Sub
        End If

         If Cel.Value = Hoje Then
         Cel.Activate
         Hora = Cel.Offset(0, 1).Value
         Hora = Format(Hora, "hh:mm:ss")
         End If
         i = i + 1
Next Cel
Range("A2").Select

Tempo = Format(Hoje & Space(1) & Hora, "dd/mm/yyyy hh:mm:ss")
MsgBox Tempo
'MsgBox Format(Tempo, "dd/mm/yyyy hh:mm:ss")


    'If MsgBox("Você deseja anexar algum arquivo?", vbQuestion + vbYesNo + vbDefaultButton2, "Diretorio do arquivo") = vbYes Then
            'Caminho = InputBox("Digite o caminho do anexo: ", "Diretorio do arquivo", "caminho")
            'Else
               'MsgBox "Email sem anexo!", vbInformation, "Importante"
               'Caminho = Empty
            'End If

Application.OnTime TimeValue(Tempo), "MandaEmail"
            
Application.StatusBar = Empty
Application.ScreenUpdating = True

End Sub
