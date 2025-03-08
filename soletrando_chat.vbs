Option Explicit

Dim palavraSorteada, palavraDigitada, i, nivel, pontuacao, pulos
Dim audio

' Inicializa o sistema de voz
Set audio = CreateObject("SAPI.SPVOICE")
audio.Rate = 2
audio.Volume = 50

' Configurações iniciais
pontuacao = 0
pulos = 3

' Função para sortear palavras de cada nível
Function sortearPalavra(nivel)
    Randomize
    Select Case nivel
        Case "A"
            Select Case Int(Rnd * 10) + 1
                Case 1: sortearPalavra = "Gato"
                Case 2: sortearPalavra = "Casa"
                Case 3: sortearPalavra = "Sol"
                Case 4: sortearPalavra = "Bola"
                Case 5: sortearPalavra = "Amigo"
                Case 6: sortearPalavra = "Livro"
                Case 7: sortearPalavra = "Carro"
                Case 8: sortearPalavra = "Cadeira"
                Case 9: sortearPalavra = "Água"
                Case 10: sortearPalavra = "Flor"
            End Select
        Case "B"
            Select Case Int(Rnd * 10) + 1
                Case 1: sortearPalavra = "Distorção"
                Case 2: sortearPalavra = "Licença"
                Case 3: sortearPalavra = "Aquiescer"
                Case 4: sortearPalavra = "Experiência"
                Case 5: sortearPalavra = "Suspense"
                Case 6: sortearPalavra = "Exclamar"
                Case 7: sortearPalavra = "Meteorologia"
                Case 8: sortearPalavra = "Reivindicar"
                Case 9: sortearPalavra = "Coincidência"
                Case 10: sortearPalavra = "Execução"
            End Select
        Case "C"
            Select Case Int(Rnd * 10) + 1
                Case 1: sortearPalavra = "Exceção"
                Case 2: sortearPalavra = "Psicologia"
                Case 3: sortearPalavra = "Apêndice"
                Case 4: sortearPalavra = "Fotossíntese"
                Case 5: sortearPalavra = "Catástrofe"
                Case 6: sortearPalavra = "Persuasão"
                Case 7: sortearPalavra = "Pneumonia"
                Case 8: sortearPalavra = "Sobrancelha"
                Case 9: sortearPalavra = "Ziguezague"
                Case 10: sortearPalavra = "Inconstitucional"
            End Select
        Case "D"
            Select Case Int(Rnd * 10) + 1
                Case 1: sortearPalavra = "Idiossincrasia"
                Case 2: sortearPalavra = "Supérfluo"
                Case 3: sortearPalavra = "Réveillon"
                Case 4: sortearPalavra = "Anticonstitucionalmente"
                Case 5: sortearPalavra = "Antropocentrismo"
                Case 6: sortearPalavra = "Impeachment"
                Case 7: sortearPalavra = "Inconstitucionalissimamente"
                Case 8: sortearPalavra = "Idiossincrasia"
                Case 9: sortearPalavra = "Incompatibilização"
                Case 10: sortearPalavra = "Fosforescência"
            End Select
    End Select
End Function

' Função para jogar um nível
Sub jogarNivel(nivel, palavrasQtd, premio)
    For i = 1 To palavrasQtd
        palavraSorteada = sortearPalavra(nivel)
        audio.Speak "A palavra sorteada é " & palavraSorteada

        palavraDigitada = InputBox("Digite a palavra sorteada ou 'pular' para usar um pulo (" & pulos & " restantes):", "Nível " & nivel)

        ' Verifica se o jogador quer pular
        If LCase(palavraDigitada) = "pular" And pulos > 0 Then
            pulos = pulos - 1
            MsgBox "Palavra pulada! Você ainda tem " & pulos & " pulos restantes.", vbInformation, "Pulo Utilizado"
        ElseIf palavraDigitada = palavraSorteada Then
            pontuacao = pontuacao + premio
        Else
            MsgBox "Palavra incorreta, você perdeu!" & vbNewLine & "A palavra era: " & palavraSorteada, vbCritical, "Fim de Jogo"
            Exit Sub
        End If
    Next
End Sub

' Rodar os níveis conforme as regras
Call jogarNivel("A", 5, 1000)     ' Nível A - 5 palavras
Call jogarNivel("B", 5, 10000)    ' Nível B - 5 palavras
Call jogarNivel("C", 5, 100000)   ' Nível C - 5 palavras
Call jogarNivel("D", 1, 1000000)  ' Nível D - 1 palavra

' Exibir pontuação final
MsgBox "Jogo finalizado! Sua pontuação total foi: R$ " & pontuacao, vbInformation, "Fim de Jogo"
