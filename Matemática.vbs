dim n1,n2,operador(3),vitoria,op,resp,resposta,resp1,audio
call carregar_vitoria

sub carregar_vitoria()
vitoria=0
call carregar_pergunta
end sub

sub carregar_pergunta()
operador(0)="+"
operador(1)="-"
operador(2)="*"
randomize(second(time))
op=int(rnd * 3)
n1=int(rnd * 10)
n2=int(rnd * 10)
if op=0 then
    resposta=n1+n2
elseif op=1 then
    resposta=n1-n2
elseif op=2 then
    resposta=n1*n2
end if
resp1=cint(inputbox("========================="+vbnewline &_
                    "                   "& n1 &" "& operador(op) &" "& n2 &""+vbnewline &_
                    "========================="+vbnewline &_
                    "DIGITE 777 PARA ENCERRAR"))
if resp1=777 then
    wscript.quit
elseif resp1=resposta then
    vitoria=vitoria+1
    msgbox("PARABENS, VOCE ACERTOU!"+vbnewline &_
            "ACERTOS: "& vitoria &""),vbinformation+vbokonly,"AVISO"
else
    msgbox("INFELIZMENTE VOCE ERROU!"+vbnewline &_
            "RESPOSTA: "& resposta &""+vbnewline &_
            "VOCE ACERTOU: "& vitoria &""),vbinformation+vbokonly,"AVISO"
    vitoria=0
end if
call carregar_pergunta
end sub