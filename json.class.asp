<%
' Classe utilizada para interpretacao e construcao de objetos JSON

class JSON
	dim i_dicionario, i_debug

	' Properties
	public property get debug
		debug = i_debug
	end property
	
	public property let debug(value)
		i_debug = value
	end property

	' Constructor and destructor
	private sub class_initialize()
		set i_dicionario = Server.CreateObject("Scripting.Dictionary")
		i_debug = false
	end sub
	
	private sub class_terminate()
		i_dicionario.removeAll
		set i_dicionario = nothing
	end sub
	
	
	' Methods
	public sub load(byval strJson)
		dim regex, i, size, char, prevchar, quoted
		dim mode, item, key, value, openArray
		dim actualLCID, curentArray, currentObject
		
		log("Load string: """ & strJson & """")
		
		actualLCID = session.LCID
		session.LCID = 1033
		
		strJson = trim(strJson)
		
		key = "[[root]]"
		i = 0
		size = len(strJson)
		
		' Se não tiver o mínimo de 2 caracteres, sai
		if size < 2 then
			load = ""
			exit sub
		end if
		
		' Inicializa o objeto regex para usar durante o loop
		set regex = new regexp
		regex.global = true
		regex.ignoreCase = true
		regex.pattern = "\w"
		
		mode = "init"
		quoted = false
		openArray = false
		currentArray = null
		set currentObject = i_dicionario
		
		do while i < size
			i = i + 1
			char = mid(strJson, i, 1)
			
			' Raiz ou início do objeto
			if mode = "init" then
				log("Enter init")
				
				' Se for o a raiz
				if key = "[[root]]" then
					' então esvazia o objeto
					i_dicionario.removeAll
				end if
				
				
				' Init object
				if char = "{" then
					log("Create object")
					
					if key <> "[[root]]" then
						' se não, cria um novo objeto
						set item = createObject("scripting.fileSystemObject")
						currentObject.add key, item
						
						item.add "__parent", currentObject
						set currentObject = item
					end if
					mode = "openKey"
					
				' Init Array
				elseif char = "[" then
					openArray = true
					log("Create array")
					
					redim item(-1)
					currentObject.add key, item
					
					currentArray = item
					
					mode = "openValue"
				end if
			
			' Iniciando uma chave
			elseif mode = "openKey" then
				key = ""
				if char = """" then
					log("Open key")
					mode = "closeKey"
				end if
			
			' Preenche a chave até encontrar uma aspa dupla
			elseif mode = "closeKey" then
				' Se encontrar, então inicia a busca por valores
				if char = """" and prevchar <> "\" then
					log("Close key")
					mode = "preValue"
				else
					key = key & char
				end if
			
			' Espera até os : para iniciar um valor
			elseif mode = "preValue" then
				if char = ":" then
					mode = "openValue"
					log("Open value for """ & key & """")
				end if
			
			' Iniciando um valor	
			elseif mode = "openValue" then
				value = ""
				
				' Se abrir aspas duplas, começa uma string
				if char = """" then
					log("Open string value")
					quoted = true
					mode = "closeValue"
				
				' Se abir [ começa um array
				elseif char = "[" then
					log("Open array value")
					quoted = false
					openArray = true
					mode = "init"
					i = i - 1
					
				else
					' Se for numero
					if regex.pattern <> "\d" then regex.pattern = "\d"
					if regex.test(char) then
						log("Open numeric value")
						quoted = false
						value = char
						mode = "closeValue"
					end if
				end if
			
			' Preenche o valor até finalizar
			elseif mode = "closeValue" then
				
				if quoted then
					
					if char = """" and prevchar <> "\" then
						log("Close string value: " & value)
						mode = "addValue"
					else
						value = value & char
					end if
				else
					' Se for numero
					if regex.pattern <> "\d" then regex.pattern = "\d"
					if regex.test(char) then
						value = value & char
					
					' Se o valor anterior foi um numero
					elseif regex.test(prevchar) then
						if char = "." or char = "e" then
							value = value & char
						else
							log("Close numeric value: " & value)
							mode = "addValue"
							i = i - 1
						end if
					else
						log("Close numeric value: " & value)
						mode = "addValue"
						i = i - 1
					end if
				end if
			
			' Adiciona o valor ao dicionario
			elseif mode = "addValue" then
				if key <> "" then
					if not quoted then
						log("Value converted to number")
						value = cdbl(value)
					end if
					
					quoted = false
					
					if openArray then
						redim preserve currentArray(ubound(currentArray) + 1)
						currentArray(ubound(currentArray)) = value
						currentObject(key) = currentArray
						
						log("Value added to array: """ & key & """: " & value)
					else
						currentObject.add key, value
						log("Value added: """ & key & """")
					end if
				end if
				
				mode = "next"
				i = i - 1
			
			' Muda o modo de acordo com o estado atual
			elseif mode = "next" then
				if char = "," then
					if openArray then
						log("New value")
						mode = "openValue"
					else
						log("New key")
						mode = "openKey"
					end if
					
				elseif char = "]" and prevchar <> "\" then
					log("Close array")
					openArray = false
				end if
			end if
			
			prevchar = char
		loop
		
		set regex = nothing
		
		session.LCID = actualLCID
	end sub
	
	' Aciciona uma propriedade ao objeto
	public sub add(byval prop, byval obj)
		if isArray(obj) then
			i_dicionario.add prop, parseArray(obj)
		else
			i_dicionario.add prop, prepareValue(obj)
		end if
	end sub
	
	' Devolve a representacao do objeto como string
	public function ToString()
		dim out, value
		out = "{"
		
		for each prop in i_dicionario.keys
			if out <> "{" then out = out & ","
			value = i_dicionario(prop)
			
			out = out & """" & prop & """:"
			
			if isArray(value) then
				out = out & parseArray(value)
			else
				out = out & prepareValue(value)
			end if
		next
		
		out = out & "}"
		
		ToString = out
	end function
	
	' Escreve direto na pagina
	public sub write()
		response.write ToString
	end sub
	
	
	' Helpers
	private function prepareValue(byval value)
		dim out
		
		select case lcase(typename(value))
			case "null", "nothing"
				out = "null"
			
			case "boolean"
				out = lcase(value)
			
			case "byte", "integer", "long", "single", "double", "currency", "decimal"
				out = value
			
			case "string", "char"
				out = """" & value & """"
			
			case else
				out = """" & typename(value) & """"
		end select
		
		prepareValue = out
	end function
	
	
	private function parseArray(byref arr)
		dim i, j, dimensions, out, elm
		
		out = "["
		
		dimensions = NumDimensions(arr)
		
		for i = 1 to dimensions
			if i > 1 then out = out & ","
			
			if dimensions > 1 then out = out & "["
			
			for j = 0 to ubound(arr, i)
				if dimensions > 1 then
					elm = arr(i - 1, j)
				else
					elm = arr(j)
				end if
				
				if j > 0 then out = out & ","
				
				if isArray(elm) then
					out = out & parseArray(elm)
				else
					out = out & prepareValue(elm)
				end if
				
			next
			if dimensions > 1 then out = out & "]"
		next
		
		out = out & "]"
		
		parseArray = out
	end function
	
	private Function NumDimensions(byref arr) 
		Dim dimensions : dimensions = 0 
		On Error Resume Next 
		Do While Err.number = 0 
			dimensions = dimensions + 1 
			UBound arr, dimensions 
		Loop 
		On Error Goto 0 
		NumDimensions = dimensions - 1 
	End Function 
	
	
	private sub log(byval msg)
		if i_debug then response.write "<li>" & msg & "</li>" & vbcrlf
	end sub
end class
%>