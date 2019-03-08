Sub MandaEmail()
    
    Dim EnviarPara As String
    Dim Mensagem As String
    
        EnviarPara = "erick.oliveira@inmetrics.com.br"
            Mensagem = "Teste"
            Envia_Emails EnviarPara, Mensagem
End Sub

Sub Envia_Emails(EnviarPara As String, Mensagem As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    With OutlookMail
        .To = EnviarPara
        .CC = ""
        .BCC = ""
        .Subject = "Subject - Email Teste"
        .Body = Mensagem
        .Send
        '.Display ' para envia o email diretamente defina o código  .Send
    End With
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

Sub Executar()

Dim Tempo As String

Tempo = Now + TimeValue("00:02:00")
MsgBox Format(Tempo, "dd/mm/yyyy hh:mm:ss")
Application.OnTime TimeValue(Tempo), "MandaEmail"

End Sub