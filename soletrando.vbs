dim nivelA(10), nivelB(10), nivelC(10), nivelD(10), palavraSorteada, palavraDigitada, i, n
dim audio, resp

' Carregar a voz
call carregar_voz

' Função para carregar a voz
function carregar_voz()
    set audio = createobject("SAPI.SPVOICE")
    audio.rate = 2
    audio.volume = 50
end function

' Função para mostrar o nível A
call nivelA
call nivelB
call nivelC
call nivelD

' Função para o nível A
function nivelA()
    nivelA(1) = "Gato"
    nivelA(2) = "Casa"
    nivelA(3) = "Sol"
    nivelA(4) = "Bola"
    nivelA(5) = "Amigo"
    nivelA(6) = "Livro"
    nivelA(7) = "Carro"
    nivelA(8) = "Cadeira"
    nivelA(9) = "Água"
    nivelA(10) = "Flor"

    i = 1
    do while i <= 5
        Randomize(Second(time))
        n = int(rnd * 10) + 1
        palavraSorteada = nivelA(n)

        call carregar_voz
        audio.speak "A palavra sorteada é " & palavraSorteada

        palavraDigitada = inputbox("Digite a palavra sorteada: ")

        if palavraDigitada <> palavraSorteada then
            msgbox("Palavra incorreta, você perdeu!" + vbnewline & _
                   "A palavra era: " & palavraSorteada, "Aviso")
            exit do ' Se errar, o jogo para
        else
            i = i + 1
        end if
    loop
end function

' Função para o nível B
function nivelB()
    nivelB(1) = "distorção"
    nivelB(2) = "licença"
    nivelB(3) = "aquiescer"
    nivelB(4) = "experiência"
    nivelB(5) = "suspense"
    nivelB(6) = "exclamar"
    nivelB(7) = "meteorologia"
    nivelB(8) = "reivindicar"
    nivelB(9) = "coincidência"
    nivelB(10) = "execução"

    i = 1
    do while i <= 5
        Randomize(Second(time))
        n = int(rnd * 10) + 1
        palavraSorteada = nivelB(n)

        call carregar_voz
        audio.speak "A palavra sorteada é " & palavraSorteada

        palavraDigitada = inputbox("Digite a palavra sorteada: ")

        if palavraDigitada <> palavraSorteada then
            msgbox("Palavra incorreta, você perdeu!" + vbnewline & _
                   "A palavra era: " & palavraSorteada, "Aviso")
            exit do ' Se errar, o jogo para
        else
            i = i + 1
        end if
    loop
end function

' Função para o nível C
function nivelC()
    nivelC(1) = "Exceção"
    nivelC(2) = "Psicologia"
    nivelC(3) = "Apêndice"
    nivelC(4) = "Fotossíntese"
    nivelC(5) = "Catástrofe"
    nivelC(6) = "Persuasão"
    nivelC(7) = "Pneumonia"
    nivelC(8) = "Sobrancelha"
    nivelC(9) = "Ziguezague"
    nivelC(10) = "inconstitucional"

    i = 1
    do while i <= 5
        Randomize(Second(time))
        n = int(rnd * 10) + 1
        palavraSorteada = nivelC(n)

        call carregar_voz
        audio.speak "A palavra sorteada é " & palavraSorteada

        palavraDigitada = inputbox("Digite a palavra sorteada: ")

        if palavraDigitada <> palavraSorteada then
            msgbox("Palavra incorreta, você perdeu!" + vbnewline & _
                   "A palavra era: " & palavraSorteada, "Aviso")
            exit do ' Se errar, o jogo para
        else
            i = i + 1
        end if
    loop
end function

' Função para o nível D
function nivelD()
    nivelD(1) = "Idiossincrasia"
    nivelD(2) = "Supérfluo"
    nivelD(3) = "réveillon"
    nivelD(4) = "anticonstitucionalmente"
    nivelD(5) = "antropocentrismo"
    nivelD(6) = "impeachment"
    nivelD(7) = "inconstitucionalissimamente"
    nivelD(8) = "idiossincrasia"
    nivelD(9) = "incompatilibização"
    nivelD(10) = "fosforescência"

    i = 1
    do while i <= 5
        Randomize(Second(time))
        n = int(rnd * 10) + 1
        palavraSorteada = nivelD(n)

        call carregar_voz
        audio.speak "A palavra sorteada é " & palavraSorteada

        palavraDigitada = inputbox("Digite a palavra sorteada: ")

        if palavraDigitada <> palavraSorteada then
            msgbox("Palavra incorreta, você perdeu!" + vbnewline & _
                   "A palavra era: " & palavraSorteada, "Aviso")
            exit do ' Se errar, o jogo para
        else
            i = i + 1
        end if
    loop
end function
