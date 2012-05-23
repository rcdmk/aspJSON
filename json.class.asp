<%
' Classe utilizada para interpretacao e construcao de objetos JSON

class JSON
	dim i_dicionario

	private sub class_initialize()
		set i_dicionario = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	private sub class_terminate()
		set i_dicionario = nothing
	end sub
	
	public sub load(byval strJson)
		dim regex, i, size, char, prevchar, quoted, mode, item, key, value, openArray
		
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
		
		mode = "root"
		openArray = false
		
		do while i < size
			i = i + 1
			char = mid(strJson, i, 1)
			
			' Raiz ou início do objeto
			if mode = "root" then
				' Init object
				if char = "{" then
					' Se for o primeiro char, é a raiz
					if i = 1 then
						' então esvazia o objeto
						i_dicionario.removeAll
					else
						' se não, cria um novo objeto
						set item = createObject("scripting.fileSystemObject")
						i_dicionario.add key, item
					end if
					
					mode = "openKey"
					
				' Init Array
				elseif char = "[" then
					openArray = true
					
					' Se for o primeiro char, é a raiz, então esvazia o objeto
					if i = 1 then
						i_dicionario.removeAll
						
						redim item(0)
						i_dicionario.add key, item
					end if
					
					mode = "openValue"
				end if
			
			' Iniciando uma chave
			elseif mode = "openKey" then
				key = ""
				
				if char = """" then  mode = "closeKey"
			
			' Preenche a chave até encontrar uma aspa dupla
			elseif mode = "closeKey" then
				' Se encontrar, então inicia a busca por valores
				if char = """" and prevchar <> "\" then
					mode = "preValue"
				else
					key = key & char
				end if
			
			' Espera até os : para iniciar um valor
			elseif mode = "preValue" then
				if char = ":" then mode = "openValue"
			
			' Iniciando um valor	
			elseif mode = "openValue" then
				value = ""
				
				' Se abrir aspas duplas, começa uma string
				if char = """" then
					quoted = true
					value = char
					mode = "closeValue"
				
				' Se abir [ começa um array
				elseif char = "[" then
					quoted = false
					mode = "openArray"
					
				else
					if regex.pattern <> "\d" then regex.pattern = "\d"
					
					if regex.test(char) then
						quoted = false
						value = char
						mode = "closeValue"
					end if
				end if
				
			elseif mode = "closeValue" then
				
				if quoted then
					value = value & char
					
					if char = """" and prevchar <> "\" then
						quoted = false
						mode = "addValue"
					end if
				else
					if regex.pattern <> "\d" then regex.pattern = "\d"
					
					if regex.test(char) then
						value = value & char
						mode = "addValue"
						
					elseif regex.test(prevchar) then
						if char = "." or char = "e" then
							value = value & char
							mode = "addValue"
						end if
					end if
				end if
			
			elseif mode = "next" then
				if char = "," then
					if openArray then
						mode = "openValue"
					else
						mode = "openKey"
					end if
				end if
				
			elseif mode = "addValue" then
				i_dicionario.add key, value
				mode = "next"	
			end if
			
			prevchar = char
		loop
		
		set regex = nothing
	end sub
	
	public sub add(byval prop, byval obj)
		if isArray(obj) then
			i_dicionario.add prop, parseArray(obj)
		else
			i_dicionario.add prop, preparValue(obj)
		end if
	end sub
	
	private function preparValue(byval value)
		dim out
		
		select case typename(value)
			case "Null"
				out = "null"
			
			case "Boolean"
				out = lcase(value)
			
			case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal"
				out = value
			
			case else
				out = """" & value & """"
		end select
		
		preparValue = out
	end function
	
	private function parseArray(byval arr)
		dim i, j, dimensions, out, elm
		
		out = "["
		
		dimensions = NumDimensions(arr)
		
		
		for i = 1 to dimensions
			if i > 1 then out = out & ", "
			
			if dimensions > 1 then out = out & "["
			
			for j = 0 to ubound(arr, i)
				if dimensions > 1 then
					elm = arr(i - 1, j)
				else
					elm = arr(j)
				end if
				
				if j > 0 then out = out & ", "
	
				
				if isArray(elm) then
					out = out & parseArray(elm)
				else
					out = out & preparValue(elm)
				end if
				
			next
			if dimensions > 1 then out = out & "]"
		next
		
		out = out & "]"
		
		parseArray = out
	end function
	
	public sub write()
		dim out
		out = "{"
		
		for each prop in i_dicionario
			if out <> "{" then out = out & ", "
		
			out = out & """" & prop & """:" & i_dicionario(prop)
		next
		
		out = out & "}"
		
		response.write out
	end sub
	
	private Function NumDimensions(arr) 
		Dim dimensions : dimensions = 0 
		On Error Resume Next 
		Do While Err.number = 0 
			dimensions = dimensions + 1 
			UBound arr, dimensions 
		Loop 
		On Error Goto 0 
		NumDimensions = dimensions - 1 
	End Function 
end class
%>