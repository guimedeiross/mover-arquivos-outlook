'Inicia Macro com o nome descrito'
Public Sub SalvarCertificadoTempera(MItem As Outlook.MailItem)
'Define a vari�vel certificado como anexo'
Dim certificado As Outlook.Attachment
'Cria uma vari�vel para destino onde ser� salvo o arquivo'
Dim destino As String
'Vari�vel para o nome do arquivo'
Dim name As String
'Var�vel para o ano de recebimento do email com o certificado'
Dim ano As String


'Verifica em cada e-mail dentro da pasta'
    For Each certificado In MItem.Attachments
        
    'Reseta vari�vel do caminho do arquivo
        destino = ""
        
    'Pega a data de recebimento do e-mail'
        ano = MItem.ReceivedTime
            
    'Separa apenas o ano da data'
        ano = Mid(ano, 7, 4)
        
    'Define o caminho onde o arquivo ser� salvo, baseado no ano de recebimento do e-mail'
        destino = "Z:\Qualidade\Certificados de Tratamento T�rmico\Certificados " & ano & "\"
        
    'Pega o nome do arquivo do certificado'
        name = certificado.DisplayName
    
    'Altera o nome para o padr�o utilizado (Certificado XXXXX)'
        name = Left(name, InStr(InStr(1, name, " ") + 1, name, " ", vbTextCompare) - 1)
    
    'Verifica caminho
        If Dir(destino) = "" Then
            'Se o caminho n�o for encontrado, ser� criado
            MsgBox ("Pasta destino do certificado n�o encontrada")
            MkDir Path:=(destino)
            MsgBox ("Caminho criado !")
        End If
           
    
    'Salva o arquivo no caminho escolhido'
        certificado.SaveAsFile destino & name & ".pdf"
        
'Vai para o pr�ximo email'
    Next

'Finaliza Macro'
End Sub
