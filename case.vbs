dim selecao, frase, resp 'declaracao de variaveis

call escolha

fuction escolha()
selecao=CInt(inputbox("[1] Verde" + vbnewline &_
                      "[2] Amarelo" + vbnewline &_
                      "[3] Vermelho" + vbnewline &_
                      "[0 ou 10] Sair", "CORES DO SEMAFORO"))

select case selecao
    case 1:
        frase="Verde - Siga"
    case 2:
        frase="Amarelo - Atenção"
    case 3:
        frase="Vermelho - Pare"
    case 0,10:
        resp=msgbox("Deseja sair?", vbQuestion+vbYesNo, "Atenção")
        if resp=vbyes Then
            WScript.Quit
        Else
            call escolha
        end If

end Function
