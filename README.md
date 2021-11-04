# mover-arquivos-outlook
Script em VBS para mover arquivos que chegam como anexo no email para alguma pasta do computador.


Para utilizar esse script, você copia todo o código, abre o outlook e pressiona "ALT+F11" então vai abrir o editor vbs do outlook, você faz as modificações de quais tipos de arquivos vc quer mover, se vai ser pdf,xml etc.. e modifica a localização da pasta e então salva o script, vc pode associar a uma regra do Outlook esse script.


Alguns Outlook, tem que habilitar a opção de executar um script, indo no regedit, no seguinte caminho:

HKEY_CURRENT_USER \Software\Microsoft\Office\16.0\Outlook\Security

Criar um novo DWORD (32 bits) com o nome de EnableUnsafeClientMailRules e coloar o valor 1.
