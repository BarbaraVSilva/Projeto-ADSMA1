Dim n1, n2, n3, media, situacao, resp, audio


call entrada_notas

Function entrada_notas()
    n1 = CDbl(inputbox("Digite a primeira nota: "))
    n2 = CDbl(inputbox("Digite a segunda nota: "))
    n3 = CDbl(inputbox("Digite a terceira nota: "))
    
    media=Round((n1+n2+n3)/3,1)
    if media < 7 then
        situacao = "Reprovado"
    else
        situacao = "Aprovado"
    end if
End Function