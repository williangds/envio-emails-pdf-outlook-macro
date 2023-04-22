Sub EnviarEmailsPDF()
    Dim appOutlook As Outlook.Application
    Dim email As Outlook.MailItem
    Dim nomeArquivo As String
    Dim caminhoPasta As String
    Dim enderecoDestinatario As String
    Dim assuntoEmail As String
    Dim corpoEmail As String
    Dim anexo As Outlook.Attachment
    
    ' Define a pasta onde os arquivos PDF estão armazenados
    caminhoPasta = "C:\Caminho\para\a\pasta"
    
    ' Cria uma instância do objeto Outlook
    Set appOutlook = New Outlook.Application
    
    ' Percorre todos os arquivos na pasta e envia um e-mail para cada arquivo PDF encontrado
    nomeArquivo = Dir(caminhoPasta & "\*.pdf")
    Do While Len(nomeArquivo) > 0
        ' Cria uma nova mensagem de e-mail
        Set email = appOutlook.CreateItem(olMailItem)
        With email
            ' Define o destinatário usando o nome do arquivo PDF
            enderecoDestinatario = nomeArquivo
            
            ' Define o assunto e o corpo do e-mail
            assuntoEmail = "Assunto do e-mail"
            corpoEmail = "Conteúdo do e-mail"
            
            ' Adiciona o anexo PDF à mensagem de e-mail
            Set anexo = .Attachments.Add(caminhoPasta & "\" & nomeArquivo, olByValue)
            
            ' Define os campos do e-mail
            .To = enderecoDestinatario
            .Subject = assuntoEmail
            .Body = corpoEmail
            
            ' Envia o e-mail
            .Send
        End With
        
        ' Libera os objetos da memória
        Set email = Nothing
        Set anexo = Nothing
        
        ' Vai para o próximo arquivo PDF
        nomeArquivo = Dir
    Loop
    
    ' Libera o objeto Outlook da memória
    Set appOutlook = Nothing
End Sub
