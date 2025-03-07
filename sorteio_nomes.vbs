dim nome(5),n,i
call sorteio
function sorteio()
nome(1)="Moquidesia"
nome(2)="Jurema"
nome(3)="Lindolfo"
nome(4)="Ademir"
nome(5)="Joselito"
'for i=1 to 10 step 1
i=1
do while i <= 10
    Randomize(Second(time)) 'Torna o sorteio dinamico baseado no SO
    n=int(rnd * 5) + 1
    msgbox(nome(n)), vbInformation+vbOKOnly, "qtde Sorteio: "& i &""
    i=i+1
Loop
'next
msgbox("fim do sorteio!"), vbInformation+vbOKOnly, "aviso"
end Function
