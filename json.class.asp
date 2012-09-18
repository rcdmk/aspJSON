<%
' Base JSON object class

class JSON
	dim i_debug, i_depth, i_parent
	dim i_properties

	' Properties
	public property get debug
		debug = i_debug
	end property
	
	public property let debug(value)
		i_debug = value
	end property
	
	
	public property get depth
		depth = i_depth
	end property
	
	private property let depth(value)
		i_depth = value
	end property
	
	
	public property get pairs
		pairs = i_properties
	end property
	
	
	public property get parent
		set parent = i_parent
	end property	
	
	public property set parent(value)
		set i_parent = value
		me.depth = i_parent.depth + 1
	end property
	
	

	' Constructor and destructor
	private sub class_initialize()
		i_depth = 0
		i_debug = false
		set i_parent = nothing
		redim i_properties(-1)
	end sub
	
	private sub class_terminate()
		dim i
		for i = 0 to ubound(i_properties)
			set i_properties(i) = nothing
		next
		
		redim i_properties(-1)
	end sub
	
	
	' Methods
	public sub parse(byval strJson)
		dim regex, i, size, char, prevchar, quoted
		dim mode, item, key, value, openArray, openObject
		dim actualLCID, tmpArray, addedToArray
		dim currentObject, currentArray
		
		log("Load string: """ & strJson & """")
		
		actualLCID = session.LCID
		session.LCID = 1033
		
		strJson = trim(strJson)
		
		i = 0
		size = len(strJson)
		
		' At least 2 chars to continue
		if size < 2 then  exit sub
		
		' Init the regex to be used in the loop
		set regex = new regexp
		regex.global = true
		regex.ignoreCase = true
		regex.pattern = "\w"
		
		key = "[[root]]"
		mode = "init"
		quoted = false
		set currentObject = me
		
		do while i < size
			i = i + 1
			char = mid(strJson, i, 1)
			
			' root or object begining
			if mode = "init" then
				log("Enter init")
				
				' if we are in root
				if key = "[[root]]" then
					' empty the object
					redim i_properties(-1)
				end if
				
				' Init object
				if char = "{" then
					log("Create object<ul>")
					
					if key <> "[[root]]" then
						' creates a new object
						set item = new JSON
						set item.parent = currentObject
						
						addedToArray = false
						
						if isArray(currentArray) then
							if currentArray.depth >= currentObject.depth then
								set item.parent = currentArray
								tmpArray = currentArray.items
								
								ArrayPush tmpArray, item
								
								currentArray.items = tmpArray
								addedToArray = true
							end if
						end if
						
						item.depth = item.parent.depth + 1
						set currentObject = item
						
						if not addedToArray then add key, item
					end if
					
					openObject = openObject + 1
					mode = "openKey"
					
				' Init Array
				elseif char = "[" then
					log("Create array<ul>")
					
					set item = new JSONarray
					
					addedToArray = false					
					
					if isobject(currentArray) and openArray > 0 then
						if currentArray.depth >= currentObject.depth then
							set item.parent = currentArray
							tmpArray = currentArray.items
							ArrayPush tmpArray, item
							
							currentArray.items = tmpArray
							addedToArray = true
						end if
					end if
					
					if not addedToArray then
						set item.parent = currentObject
						
						tmpArray = currentObject.pairs
						currentObject.value.add key, item
						item.depth = currentObject.depth + 1
					end if
					
					item.depth = item.parent.depth + 1
					
					set currentArray = item
					
					openArray = openArray + 1
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
					log("Close key: """ & key & """")
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
					mode = "init"
					i = i - 1
				
				' Se abir [ começa um array
				elseif char = "{" then
					log("Open object value")
					quoted = false
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
					dim useArray
					useArray = false
					
					if not quoted then
						log("Value converted to number")
						value = cdbl(value)
					end if
					
					quoted = false
					
					if openArray > 0 and isObject(currentArray) then
						useArray = true
						
						if isObject(currentObject) then
							if isObject(currentObject.parent) then
								if isArray(currentObject.parent.value) then useArray = false
							end if
						end if
						
						if useArray then
							tmpArray = currentArray.value
							ArrayPush tmpArray, value
							
							currentArray.value = tmpArray
							
							log("Value added to array: """ & key & """: " & value)
						end if
					end if
					
					if not useArray then
						currentObject.add key, value
						log("Value added: """ & key & """")
					end if
				end if
				
				mode = "next"
				i = i - 1
			
			' Muda o modo de acordo com o estado atual
			elseif mode = "next" then
				if char = "," then
					if openArray > 0 and isObject(currentArray) then
						if currentArray.depth >= currentObject.depth then
							log("New value")
							mode = "openValue"
						else
							log("New key")
							mode = "openKey"
						end if
					else
						log("New key")
						mode = "openKey"
					end if
					
				elseif char = "]" then
					log("Close array</ul>")
					
					if isobject(currentArray.parent) then
						set currentArray = currentArray.parent
					else
						currentArray = currentArray.parent
					end if
					
					openArray = openArray - 1
					
					mode = "next"

				elseif char = "}" then
					log("Close object</ul>")
					
					if isobject(currentObject.parent) then
						set currentObject = currentObject.parent
					else
						currentObject = currentObject.parent
					end if
					
					openObject = openObject - 1
					
					mode = "next"					
				end if
			end if
			
			prevchar = char
		loop
		
		set regex = nothing
		
		session.LCID = actualLCID
	end sub
	
	' Aciciona uma propriedade ao objeto
	public sub add(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if isObject(p) then
			err.raise 1, "A property already exists with the name: " & prop & "."
		else
			dim item
			set item = new JSONpair
			item.name = prop
			set item.parent = me
			
			if isArray(obj) then
				dim item2
				set item2 = new JSONarray
				item2.items = obj
				set item.value = item2
				
			elseif isObject(obj) then
				set item.value = obj
			else
				item.value = obj
			end if

			ArrayPush i_properties, item
		end if
	end sub
	
	' Retorna o valor da propriedade
	public function value(byval prop)
		dim p
		getProperty prop, p
		
		if isObject(p) then
			if isObject(p.value) then
				set value = p.value
			else
				value = p.value
			end if
		else
			err.raise 1, "Property " & prop & " doesn't exists."
		end if
	end function
	
	' Altera uma propriedade do objeto
	' Cria a propriedade se ela nao existir
	public sub change(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if isObject(p) then
			if isArray(obj) then
				set item = new JSONarray
				item.items = obj
				item.parent = me
				
				p.value = item
				
			elseif isObject(obj) then
				set p.value = obj
			else
				p.value = obj
			end if
		else
			add prop, obj
		end if
	end sub
	
	' Retorna a propriedade se existir
	private sub getProperty(byval prop, byref outProp)
		dim i, p
		outProp = null
		
		do while i <= ubound(i_properties)
			set p = i_properties(i)
			
			if p.name = prop then
				set outProp = p
				
				exit do
			end if
			
			i = i + 1
		loop
	end sub
	
	
	' Devolve a representacao do objeto como string
	public function Serialize()
		dim actualLCID, out
		actualLCID = session.LCID
		session.LCID = 1033
		
		out = serializeObject(me)
		
		session.LCID = actualLCID
		
		Serialize = out
	end function
	
	' Escreve direto na pagina
	public sub write()
		response.write Serialize
	end sub
	
	
	' Helpers
	private function serializeValue(byval value)
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
		
		serializeValue = out
	end function
	
	
	private function serializeArray(byref arr)
		dim i, j, dimensions, out, arr2, elm, val
		
		out = "["
		
		if isobject(arr) then
			arr2 = arr.items
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
					if isobject(elm.value) then
						set val = elm.value
					else
						val = elm.value
					end if
				else
					val = elm
				end if
				
				if isArray(val) or typeName(val) = "JSONarray" then
					out = out & serializeArray(val)
				elseif isObject(val) then
					out = out & serializeObject(val)
				else
					out = out & serializeValue(val)
				end if
				
			next
			if dimensions > 1 then out = out & "]"
		next
		
		out = out & "]"
		
		serializeArray = out
	end function
	
	
	private function serializeObject(obj)
		dim out, prop, value, i, pairs
		out = "{"
		
		pairs = obj.pairs
		
		for i = 0 to ubound(pairs)
			set prop = pairs(i)
			
			if out <> "{" then out = out & ","
			
			if isobject(prop.value) then
				set value = prop.value
			else
				value = prop.value
			end if
			
			if prop.name <> "[[root]]" then out = out & """" & prop.name & """:"
			
			if isArray(value) or typeName(value) = "JSONarray" then
				out = out & serializeArray(value)
				
			elseif isObject(value) then
				out = out & serializeObject(value)
				
			else
				out = out & serializeValue(value)
			end if
		next
		
		out = out & "}"
		
		serializeObject = out
	end function
	
	' 
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


class JSONarray
	dim i_items, i_depth, i_parent

	public property get items
		items = i_items
	end property	
	
	public property let items(value)
		if isArray(value) then
			i_items = value
		else
			err.raise 1, "The value assigned is not an array."
		end if
	end property	
	
	
	public property get depth
		depth = i_depth
	end property
	
	private property let depth(value)
		i_depth = value
	end property
	
	
	public property get parent
		set parent = i_parent
	end property	
	
	public property set parent(value)
		set i_parent = value
		me.depth = i_parent.depth + 1
	end property
	
	
	private sub class_initialize
		redim i_items(-1)
		depth = 0
	end sub
	
	private sub class_terminate
		dim i
		for i = 0 to ubound(i_items)
			set i_items(i) = nothing
		next
	end sub
end class


class JSONpair
	dim i_name, i_value
	dim i_parent
	
	
	public property get name
		name = i_name
	end property
	
	public property let name(val)
		i_name = val
	end property
	
	
	public property get value
		if isObject(i_value) then
			set value = i_value
		else
			value = i_value
		end if
	end property
	
	public property let value(val)
		i_value = val
	end property
	
	public property set value(val)
		set i_value = val
	end property
	
	
	public property get parent
		set parent = i_parent
	end property	
	
	public property set parent(val)
		set i_parent = val
	end property
	
	private sub class_initialize
	end sub
	
	private sub class_terminate
		if isObject(value) then set value = nothing
	end sub
end class
%>