'1 parte entrar com numero dos ips 
'2 gerar o ping de todos, gerar arquivo temporario, 4 Gerar o CSV 
'5 pegar o arq 2 mandar para HTML 
'salvar em csv e em HTML

dim nome, nome2 
dim fso, fso2, fso3, fso4
dim aux, aux2
dim comando, temp
dim y, x
dim linha, linha2
dim resultado, salva, final
dim dataAtual, dataFormatada

set fso=Createobject ("scripting.filesystemobject") 
set aux=CreateObject("WScript.Shell")
set aux2=CreateObject("WScript.Shell")
set fso3=CreateObject("scripting.filesystemobject")
set final=fso3.Createtextfile("log.txt",3, true)
set final1=fso3.Createtextfile("log.csv",3, true)
set final2=fso3.Createtextfile("log.html",3, true)

set fso4 = CreateObject("Scripting.FileSystemObject")
set cria= fso4.CreateTextFile("IPS.txt", True)

do while contador <1
	ips=inputbox("entre com o numero de IP")
    cria.WriteLine(IPS)
    contador=inputbox("1 encerra, 0 continua ")
loop

nome= "ips.txt"  
if fso.fileexists(nome) then  'fileexists verifica se o arquivo existe
	set pings=fso.opentextFile ("pings.txt",1,true)'2 grava por cima 3 continua o arquivo 
    set ips=fso.opentextFile(nome,1,true)'1 é leitura 
    conta=0 
    do while ips.atendofstream=false'faça isso enquanto tiver linha no arquivo 
		conta=conta+1 
        linha=ips.readline'pega uma linha do arquivo
		comando = ("cmd /k cd & ping "+linha+">pings.txt & cd & exit")
		aux.run comando
		Wscript.sleep 10000 'salavador da patria vulgo temporizador
		
		' se linha 11 do arquivo pings.txt exite então faça se não ler arquivo novamente
		set fso2 = createobject ("scripting.filesystemobject") 

		nome2= "pings.txt"  
		if fso2.fileexists(nome2) then  'fileexists verifica se o arquivo existe
			set pings=fso2.opentextFile(nome2,1,true)'1 é leitura 
			conta=0 
	
			do while pings.atendofstream=false'faça isso enquanto tiver linha no arquivo 
        		conta=conta+1 
				linha2=pings.readline'pega uma linha do arquivo 
				if conta=12 then
					dataAtual = now ( )
					dataFormatada = FormatDateTime(dataAtual, 3)
					'0 – Retorna o formato short date (caso seja passado apenas a hora o retorno será a hora, se passar a data o retorno será a data, se passar ambos, como short format)
 					'1 – long date
					'2 – short date format especificado nas Configurações Regionais do computador
					'3 – Retorna a hora especificada nas Configurações Regionais do computador
					'4 – Retorna a hora usando o formato 24horas (hh:mm)		

					'msgbox(linha2)
					x=split(linha2,"ms")			
					y=split(x(2), "=")
					resultado=(y(1))

					if Resultado>0 and Resultado <=50 then		
						'chama arq html
						salva=("ip Testado:"+linha+" Media "+resultado+"Otimo Hora: "+dataFormatada)
						final.WriteLine(salva)
						final1.WriteLine(salva)
        	    		final2.WriteLine("<font size='3' color='red'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")
					end if

					if Resultado >=50 and Resultado <=130 then
						final.WriteLine(salva)
						final1.WriteLine(salva)
						final2.WriteLine("<font size='3' color='gren'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")			
						salva=("ip Testado:"+linha+" Media "+resultado+"BOM Hora: "+dataFormatada)
					end if

					if Resultado >=131 and Resultado <=260 then
						final.WriteLine(salva)
						final1.WriteLine(salva)
						final2.WriteLine("<font size='3' color='red'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")			
						salva=("ip Testado:"+linha+" Media "+resultado+"REGULAR Hora: "+dataFormatada)
					end if

					if Resultado >=261 and Resultado <=260 then
						final.WriteLine(salva)
						final1.WriteLine(salva)
						final2.WriteLine("<font size='3' color='blue'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")
						salva=("ip Testado:"+linha+" Media "+resultado+"ACEITAVEL Hora: "+dataFormatada)
					end if

					if Resultado >=261 and Resultado <=500 then
						final.WriteLine(salva)
						final1.WriteLine(salva)
        	    		final2.WriteLine("<font size='3' color='yellow'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")			
						salva=("ip Testado:"+linha+" Media "+resultado+"RUIM Hora: "+dataFormatada)
					end if

					if Resultado >=501 then
						final.WriteLine(salva)
						final1.WriteLine(salva)
        	    		final2.WriteLine("<font size='3' color='pink'>")
						final2.WriteLine(salva)
        	    		final2.Writeline("</font></br>")	
						salva=("ip Testado:"+linha+" Media "+resultado+"DEU MERDA REINICIA O SERVIÇO0 Hora: "+dataFormatada)		
					end if		
				end if
    		loop 
		end if			
	loop
else
	msgbox("arquivo inexistente") 
end if

msgbox ("Programa terminado")
