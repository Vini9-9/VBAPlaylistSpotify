Attribute VB_Name = "CriarPlaylistSpotify"
Public Sub CriarPlaylistSpotify()

Dim driver As New ChromeDriver
Dim keys As New Selenium.keys
Dim oCheck As New By ' Objeto será usado no método.
Const sBtnPath As String = "/html/body/div[4]/div/div[2]/div[3]/main/div[2]/div[2]/div/div/div[2]/section/div[2]/div[3]/div/div/div/div[2]/div[13]"
Const sPlaylistBtnPath As String = "/html/body/div[4]/div/div[2]/nav/div[1]/div[2]/div/div[1]/button"
Const sInputBtnPath As String = "/html/body/div[4]/div/div[2]/div[3]/main/div[2]/div[2]/div/div/div[2]/section/div[2]/div[3]/section/div/div/input"

On Error GoTo MsgError

lastrow = Range("a3").End(xlDown).Row

If MsgBox("Criar playlist com " & lastrow - 3 & " músicas?", vbYesNo) = vbNo Then Exit Sub

With driver
    .Start
    
'Tela inicial Spotify

    .Get "https://accounts.spotify.com/pt-BR/login?continue=https:%2F%2Fopen.spotify.com%2F"
    
'Tela de Login
    
    'nome do usuário
    'id = "login-username"
    
    .FindElementById("login-username").Click
    .SendKeys (Trim(InputBox("informe o e-mail")))
    
    'senha
    'id="login-password"
    
    .FindElementById("login-password").Click
    .SendKeys (Trim(InputBox("informe a senha")))
    
    'botão Login
    'id="login-button"

    .FindElementById("login-button").Click
    
'Tela inicial logada
    
    .Wait (1000)
    
    'botão Criar Playlist
    
tentarNovamente:
    
    .FindElementByXPath("/html/body/div[4]/div/div[2]/nav/div[1]/div[2]/div/div[1]/button", 2000).Click
    
    .Wait (1000)
    
    'Caso não encontre o proximo ele clica novamente no botão de criar playlist
    
    If .IsElementPresent(oCheck.XPath(sInputBtnPath)) Then
        Set hFrame = .FindElementByXPath(sInputBtnPath)
        hFrame.Click
    Else
    
        GoTo tentarNovamente
    
    End If

For i = 4 To lastrow
    
    .Wait (1000)
    
    .SendKeys (Range("A" & i).Value)
    
    .Wait (1000)
    
    'Ver todas as músicas
    
    If .IsElementPresent(oCheck.XPath(sBtnPath)) Then
        Set hFrame = .FindElementByXPath(sBtnPath, 3)
        hFrame.Click
    Else
    
        Range("B" & i).Value = "Não localizado"
    
        GoTo proximo
    
    End If
   
   'Adicionar música
   
    .FindElementByXPath("/html/body/div[4]/div/div[2]/div[3]/main/div[2]/div[2]/div/div/div[2]/section/div[2]/div[3]/div/div/div/div[2]/div[1]/div/div[3]/button/span").Click
    
    .Wait (1000)
    
limpar:
    
    If .FindElementByXPath(sInputBtnPath).Value <> "" Then
        Set hFrame = .FindElementByXPath(sInputBtnPath)
        hFrame.Clear
        hFrame.Click
    End If

    Range("B" & i).Value = "OK"

proximo:
    
Next i
    
'Tela de detalhes da playlist
    
    .FindElementByXPath("/html/body/div[4]/div/div[2]/div[3]/main/div[2]/div[2]/div/div/div[2]/section/div[1]/div[5]/span/button/span/h1").Click
    
    'Nome da Playlist
    
    .FindElementByXPath("/html/body/div[15]/div/div/div/div[2]/div[2]/input").Clear
    .FindElementByXPath("/html/body/div[15]/div/div/div/div[2]/div[2]/input").Click
    .SendKeys (Range("B1").Value)
    
    'Descrição da Playlist
    
    .FindElementByXPath("/html/body/div[15]/div/div/div/div[2]/div[3]/textarea").Click
    .SendKeys (Range("B2").Value)
    
    'Salvar
    
    .FindElementByXPath("/html/body/div[15]/div/div/div/div[2]/button").Click
    
    'Compartilhar link da playlist
    
    .FindElementByXPath("/html/body/div[4]/div/div[2]/div[3]/main/div[2]/div[2]/div/div/div[2]/section/div[2]/div[2]/div/button[2]").Click
    
    .FindElementByXPath("/html/body/div[13]/div/ul/li[6]/button").Click
    
    .FindElementByXPath("/html/body/div[13]/div/ul/li[6]/div/ul/li[1]/button").Click
    
    url_spotify = .GetClipBoard
    
    Range("c1").Value = url_spotify
    
End With

MsgBox ("Playlist criada com sucesso :)")

Exit Sub

MsgError:

MsgBox ("Ocoreu um erro inesperado :(")

End Sub
