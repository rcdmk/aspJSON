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
		dim tmpArray, tmpObj
		
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
		set currentObject = new JSONItem
		set currentObject.value = i_dicionario
		
		do while i < size
			i = i + 1
			char = mid(strJson, i, 1)
			
			' Raiz ou início do objeto
			if mode = "init" then
				log("Enter init")
				
				' Se for o a raiz
				if key = "[[root]]" then
					' então esvazia o objeto
					currentObject.value.removeAll
				end if
				
				
				' Init object
				if char = "{" then
					log("Create object")
					
					if key <> "[[root]]" then
						' se não, cria um novo objeto
						set item = new JSONitem
						set item.parent = currentObject
						set item.value = createObject("scripting.fileSystemObject")
						
						currentObject.add key, item
						set currentObject = item
					end if
					mode = "openKey"
					
				' Init Array
				elseif char = "[" then
					log("Create array")
					
					redim item1(-1)
					
					set item = new JSONitem
					item.value = item1
					
					if isobject(currentArray) and openArray then
						set item.parent = currentArray
						tmpArray = currentArray.value
						ArrayPush tmpArray, item
						
						currentArray.value = tmpArray
					else
						currentObject.value.add key, item
					end if
					
					set currentArray = item
					
					openArray = true
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
						log("Close string value: """ & value & """")
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
						tmpArray = currentArray.value
						ArrayPush tmpArray, value
						
						currentArray.value = tmpArray
						
						log("Value added to array: """ & key & """: " & value)
					else
						currentObject.value.add key, value
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
					
				elseif char = "]" then
					log("Close array")
					
					if isobject(currentArray.parent) then
						set currentArray = currentArray.parent
					else
						openArray = false
					end if

				elseif char = "}" then
					log("Close object")
					
					if isobject(currentObject.parent) then
						set currentObject = currentObject.parent
					end if
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
			
			if isobject(i_dicionario(prop)) then
				value = i_dicionario(prop).value
			else
				value = i_dicionario(prop)
			end if
			
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
		dim i, j, dimensions, out, arr2, elm, val
		
		out = "["
		
		if isobject(arr) then
			arr2 = arr.value
		else
			arr2 = arr
		end if
		
		dimensions = NumDimensions(arr2)
		
		for i = 1 to dimensions
			if i > 1 then out = out & ","
			
			if dimensions > 1 then out = out & "["
			
			for j = 0 to ubound(arr2, i)
				if dimensions > 1 then
					if isobject(arr2(i - 1, j)) then
						set elm = arr2(i - 1, j)
					else
						elm = arr2(i - 1, j)
					end if
				else
					if isobject(arr2(j)) then
						set elm = arr2(j)
					else
						elm = arr2(j)
					end if
				end if
				
				if j > 0 then out = out & ","
				
				if isobject(elm) then
					val = elm.value
				else
					val = elm
				end if
				
				if isArray(val) then
					out = out & parseArray(val)
				else
					out = out & prepareValue(val)
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
	
	private function ArrayPush(byref arr, byref value)
		redim preserve arr(ubound(arr) + 1)
		if isobject(value) then
			set arr(ubound(arr)) = value
		else
			arr(ubound(arr)) = value
		end if
		ArrayPush = arr
	end function
	
	private sub log(byval msg)
		if i_debug then response.write "<li>" & msg & "</li>" & vbcrlf
	end sub
end class


class JSONitem
	dim i_value
	
	public parent
	
	public property get value
		if isObject(i_value) then
			set value = i_value
		else
			value = i_value
		end if
	end property
	
	public property set value(vl)
		set i_value = vl
	end property
	
	public property let value(vl)
		i_value = vl
	end property
	
	private sub class_terminate
		set i_value = nothing
	end sub
end class
%>