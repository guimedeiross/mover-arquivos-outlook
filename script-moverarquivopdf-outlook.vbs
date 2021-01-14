'Inicia Macro com o nome descrito'
Public Sub SalvarCertificadoTempera(MItem As Outlook.MailItem)
'Define a variável certificado como anexo'
Dim certificado As Outlook.Attachment
'Cria uma variável para destino onde será salvo o arquivo'
Dim destino As String
'Variável para o nome do arquivo'
Dim name As String
'Varável para o ano de recebimento do email com o certificado'
Dim ano As String


'Verifica em cada e-mail dentro da pasta'
    For Each certificado In MItem.Attachments
        
    'Reseta variável do caminho do arquivo
        destino = ""
        
    'Pega a data de recebimento do e-mail'
        ano = MItem.ReceivedTime
            
    'Separa apenas o ano da data'
        ano = Mid(ano, 7, 4)
        
    'Define o caminho onde o arquivo será salvo, baseado no ano de recebimento do e-mail'
        destino = "Z:\Qualidade\Certificados de Tratamento Térmico\Certificados " & ano & "\"
        
    'Pega o nome do arquivo do certificado'
        name = certificado.DisplayName
    
    'Altera o nome para o padrão utilizado (Certificado XXXXX)'
        name = Left(name, InStr(InStr(1, name, " ") + 1, name, " ", vbTextCompare) - 1)
    
    'Verifica caminho
        If Dir(destino) = "" Then
            'Se o caminho não for encontrado, será criado
            MsgBox ("Pasta destino do certificado não encontrada")
            MkDir Path:=(destino)
            MsgBox ("Caminho criado !")
        End If
           
    
    'Salva o arquivo no caminho escolhido'
        certificado.SaveAsFile destino & name & ".pdf"
        
'Vai para o próximo email'
    Next

'Finaliza Macro'
End Sub
