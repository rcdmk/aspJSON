<%
' JSON object class 3.5.5 - May, 29th - 2016
'
' Licence:
' The MIT License (MIT)
' Copyright (c) 2016 RCDMK - rcdmk[at]hotmail[dot]com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
' associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
' NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
' SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

const JSON_ROOT_KEY = "[[JSONroot]]"
const JSON_DEFAULT_PROPERTY_NAME = "data"
const JSON_SPECIAL_VALUES_REGEX = "^(?:(?:t(?:r(?:ue?)?)?)|(?:f(?:a(?:l(?:se?)?)?)?)|(?:n(?:u(?:ll?)?))|(?:u(?:n(?:d(?:e(?:f(?:i(?:n(?:ed?)?)?)?)?)?)?)?))$"

const JSON_ERROR_PARSE = 1
const JSON_ERROR_PROPERTY_ALREADY_EXISTS = 2
const JSON_ERROR_PROPERTY_DOES_NOT_EXISTS = 3 ' DEPRECATED
const JSON_ERROR_NOT_AN_ARRAY = 4
const JSON_ERROR_INDEX_OUT_OF_BOUNDS = 9 ' Numbered to have the same error number as the default "Subscript out of range" exeption

class JSONobject
	dim i_debug, i_depth, i_parent
	dim i_properties, i_version, i_defaultPropertyName

	' Set to true to show the internals of the parsing mecanism
	public property get debug
		debug = i_debug
	end property
	
	public property let debug(value)
		i_debug = value
	end property

	
	' Gets/sets the default property name generated when loading recordsets and arrays (default "data")
	public property get defaultPropertyName
		defaultPropertyName = i_defaultPropertyName
	end property

	public property let defaultPropertyName(value)
		i_defaultPropertyName = value
	end property


	' The depth of the object in the chain, starting with 1
	public property get depth
		depth = i_depth
	end property
	
	
	' The property pairs ("name": "value" - pairs)
	public property get pairs
		pairs = i_properties
	end property
	
	
	' The parent object
	public property get parent
		set parent = i_parent
	end property
	
	public property set parent(value)
		set i_parent = value
		i_depth = i_parent.depth + 1
	end property
	
	

	' Constructor and destructor
	private sub class_initialize()
		i_version = "3.5.5"
		i_depth = 0
		i_debug = false
		i_defaultPropertyName = JSON_DEFAULT_PROPERTY_NAME
		
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
	
	
	' Parse a JSON string and populate the object
	public function parse(byval strJson)
		dim regex, i, size, char, prevchar, quoted
		dim mode, item, key, value, openArray, openObject
		dim actualLCID, tmpArray, tmpObj, addedToArray
		dim root, currentObject, currentArray
		
		log("Load string: """ & strJson & """")
		
		' Store the actual LCID and use the en-US to conform with the JSON standard
		actualLCID = Response.LCID
		Response.LCID = 1033
		
		strJson = trim(strJson)
		
		size = len(strJson)
		
		' At least 2 chars to continue
		if size < 2 then err.raise JSON_ERROR_PARSE, TypeName(me), "Invalid JSON string to parse"
		
		' Init the regex to be used in the loop
		set regex = new regexp
		regex.global = true
		regex.ignoreCase = true
		regex.pattern = "\w"
		
		' setup initial values
		i = 0
		set root = me
		key = JSON_ROOT_KEY
		mode = "init"
		quoted = false
		set currentObject = root
		
		' main state machine
		do while i < size
			i = i + 1
			char = mid(strJson, i, 1)
			
			' root, object or array start
			if mode = "init" then
				log("Enter init")
				
				' if we are in root, clear previous object properties
				if key = JSON_ROOT_KEY and TypeName(currentArray) <> "JSONarray" then redim i_properties(-1)
				
				' Init object
				if char = "{" then
					log("Create object<ul>")
					
					if key <> JSON_ROOT_KEY or TypeName(root) = "JSONarray" then
						' creates a new object
						set item = new JSONobject
						set item.parent = currentObject
						
						addedToArray = false
						
						' Object is inside an array
						if TypeName(currentArray) = "JSONarray" then
							if currentArray.depth > currentObject.depth then
								' Add it to the array
								set item.parent = currentArray
								currentArray.Push item
								
								addedToArray = true

								log("Added to the array")
							end if
						end if
						
						if not addedToArray then
							currentObject.add key, item
							log("Added to parent object: """ & key & """")
						end if
												
						set currentObject = item
					end if
					
					openObject = openObject + 1
					mode = "openKey"
					
				' Init Array
				elseif char = "[" then
					log("Create array<ul>")
					
					set item = new JSONarray
					
					addedToArray = false
					
					' Array is inside an array
					if isobject(currentArray) and openArray > 0 then
						if currentArray.depth > currentObject.depth then
							' Add it to the array
							set item.parent = currentArray
							currentArray.Push item
							
							addedToArray = true
							
							log("Added to parent array")
						end if
					end if
					
					if not addedToArray then
						set item.parent = currentObject
						currentObject.add key, item
						log("Added to parent object")
					end if

					if key = JSON_ROOT_KEY and item.depth = 1 then
						set root = item
						log("Set as root")
					end if
					
					set currentArray = item
					openArray = openArray + 1
					mode = "openValue"
				end if
			
			' Init a key
			elseif mode = "openKey" then
				key = ""
				if char = """" then
					log("Open key")
					mode = "closeKey"
				elseif char = "}" then ' empty objects
					log("Empty object")
					mode = "next"
					i = i - 1 ' we backup one char to make the next iteration get the closing bracket
				end if
			
			' Fill in the key until finding a double quote "
			elseif mode = "closeKey" then
				' If it finds a non scaped quotation, change to value mode
				if char = """" and prevchar <> "\" then
					log("Close key: """ & key & """")
					mode = "preValue"
				else
					key = key & char
				end if
			
			' Wait until a colon char (:) to begin the value
			elseif mode = "preValue" then
				if char = ":" then
					mode = "openValue"
					log("Open value for """ & key & """")
				end if
			
			' Begining of value
			elseif mode = "openValue" then
				value = ""
				
				' If the next char is a closing square barcket (]), its closing an empty array
				if char = "]" then
					log("Closing empty array")
					quoted = false
					mode = "next"
					i = i - 1 ' we backup one char to make the next iteration get the closing bracket
				
				' If it begins with a double quote, its a string value
				elseif char = """" then
					log("Open string value")
					quoted = true
					mode = "closeValue"
				
				' If it begins with open square bracket ([), its an array
				elseif char = "[" then
					log("Open array value")
					quoted = false
					mode = "init"
					i = i - 1 ' we backup one char to init with '['
				
				' If it begins with open a bracket ({), its an object
				elseif char = "{" then
					log("Open object value")
					quoted = false
					mode = "init"
					i = i - 1 ' we backup one char to init with '{'
					
				else
					' If its a number, start a numeric value
					if regex.pattern <> "\d" then regex.pattern = "\d"
					if regex.test(char) then
						log("Open numeric value")
						quoted = false
						value = char
						mode = "closeValue"
						if prevchar = "-" then
							value = prevchar & char
						end if
						
					' special values: null, true, false and undefined
					elseif char = "n" or char = "t" or char = "f" or char = "u" then
						log("Open special value")
						quoted = false
						value = char
						mode = "closeValue"
					end if
				end if
			
			' Fill in the value until finish
			elseif mode = "closeValue" then
				if quoted then
					if char = """" and prevchar <> "\" then
						log("Close string value: """ & value & """")
						mode = "addValue"
						
					' special and escaped chars
					elseif prevchar = "\" then
						select case char
							case "n"
								value = value & vblf
							case "r"
								value = value & vbcr
							case "t"
								value = value & vbtab
							case else
								value = value & char
						end select
					elseif char <> "\" then
						value = value & char
					end if
				else
					' possible boolean and null values
					if regex.pattern <> JSON_SPECIAL_VALUES_REGEX then regex.pattern = JSON_SPECIAL_VALUES_REGEX
					if regex.test(char) or regex.test(value) then
						value = value & char
						if value = "true" or value = "false" or value = "null" or value = "undefined" then mode = "addValue"
					else
						char = lcase(char)
						
						' If is a numeric char
						if regex.pattern <> "\d" then regex.pattern = "\d"
						if regex.test(char) then
							value = value & char
						
						' If it's not a numeric char, but the prev char was a number
						' used to catch separators and special numeric chars
						elseif regex.test(prevchar) or prevchar = "e" then
							if char = "." or char = "e" or (prevchar = "e" and (char = "-" or char = "+")) then
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
				end if
			
			' Add the value to the object or array
			elseif mode = "addValue" then
				if key <> "" then
					dim useArray
					useArray = false
					
					if not quoted then
						if isNumeric(value) then
							log("Value converted to number")
							value = cdbl(value)
						else
							log("Value converted to " & value)
							value = eval(value)
						end if
					end if
					
					quoted = false
					
					' If it's inside an array
					if openArray > 0 and isObject(currentArray) then
						useArray = true
						
						' If it's a property of an object that is inside the array
						' we add it to the object instead
						if isObject(currentObject) then
							if currentObject.depth >= currentArray.depth + 1 then useArray = false
						end if
						
						' else, we add it to the array
						if useArray then
							tmpArray = currentArray.items
							ArrayPush tmpArray, value
							
							currentArray.items = tmpArray
							
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
			
			' Change the current mode according to the current state
			elseif mode = "next" then
				if char = "," then
					' If it's an array
					if openArray > 0 and isObject(currentArray) then
						' and the current object is a parent or sibling object
						if currentArray.depth >= currentObject.depth then
							' start an array value
							log("New value")
							mode = "openValue"
						else
							' start an object key
							log("New key")
							mode = "openKey"
						end if
					else
						' start an object key
						log("New key")
						mode = "openKey"
					end if
				
				elseif char = "]" then
					log("Close array</ul>")
					
					' If it's and open array, we close it and set the current array as its parent
					if isobject(currentArray.parent) then
						if TypeName(currentArray.parent) = "JSONarray" then
							set currentArray = currentArray.parent
						
						' if the parent is an object
						elseif TypeName(currentArray.parent) = "JSONobject" then
							set tmpObj = currentArray.parent
							
							' we search for the next parent array to set the current array
							while isObject(tmpObj) and TypeName(tmpObj) = "JSONobject"
								if isObject(tmpObj.parent) then
									set tmpObj = tmpObj.parent
								else
									tmpObj = tmpObj.parent
								end if
							wend
							
							set currentArray = tmpObj
						end if
					else
						currentArray = currentArray.parent
					end if
					
					openArray = openArray - 1
					
					mode = "next"

				elseif char = "}" then
					log("Close object</ul>")
					
					' If it's an open object, we close it and set the current object as it's parent
					if isobject(currentObject.parent) then
						if TypeName(currentObject.parent) = "JSONobject" then
							set currentObject = currentObject.parent
						
						' If the parent is and array
						elseif TypeName(currentObject.parent) = "JSONarray" then
							set tmpObj = currentObject.parent
							
							' we search for the next parent object to set the current object
							while isObject(tmpObj) and TypeName(tmpObj) = "JSONarray"
								set tmpObj = tmpObj.parent
							wend
							
							set currentObject = tmpObj
						end if
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
		
		Response.LCID = actualLCID
		
		set parse = root
	end function
	
	' Add a new property (key-value pair)
	public sub add(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if GetTypeName(p) = "JSONpair" then
			err.raise JSON_ERROR_PROPERTY_ALREADY_EXISTS, TypeName(me), "A property already exists with the name: " & prop & "."
		else
			dim item
			set item = new JSONpair
			item.name = prop
			set item.parent = me

			dim itemType
			itemType = GetTypeName(obj)

			if isArray(obj) then
				dim item2
				set item2 = new JSONarray
				item2.items = obj
				set item2.parent = me

				set item.value = item2
				
			elseif itemType = "Field" then
				item.value = obj.value
			elseif isObject(obj) and itemType <> "IStringList" then
				set item.value = obj
			else
				item.value = obj
			end if

			ArrayPush i_properties, item
		end if
	end sub
	
	' Remove a property from the object (key-value pair)
	public sub remove(byval prop)
		dim p, i
		i = getProperty(prop, p)
		
		' property exists
		if i > -1 then ArraySlice i_properties, i
	end sub
	
	' Return the value of a property by its key
	public default function value(byval prop)
		dim p
		getProperty prop, p
		
		if GetTypeName(p) = "JSONpair" then
			if isObject(p.value) then
				set value = p.value
			else
				value = p.value
			end if
		else
			value = null
		end if
	end function
	
	' Change the value of a property
	' Creates the property if it didn't exists
	public sub change(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if GetTypeName(p) = "JSONpair" then
			if isArray(obj) then
				set item = new JSONarray
				item.items = obj
				set item.parent = me
				
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
	
	' Returns the index of a property if it exists, else -1
	' @param prop as string - the property name
	' @param out outProp as variant - will be filled with the property value, nothing if not found
	private function getProperty(byval prop, byref outProp)
		dim i, p, found
		set outProp = nothing
		found = false		
		
		i = 0

		do while i <= ubound(i_properties)
			set p = i_properties(i)
			
			if p.name = prop then
				set outProp = p
				found = true
				
				exit do
			end if
			
			i = i + 1
		loop
		
		if not found then i = -1
		
		getProperty = i
	end function
	
	
	' Serialize the current object to a JSON formatted string
	public function Serialize()
		dim actualLCID, out
		actualLCID = Response.LCID
		Response.LCID = 1033
		
		out = serializeObject(me)
		
		Response.LCID = actualLCID
		
		Serialize = out
	end function
	
	' Writes the JSON serialized object to the response
	public sub write()
		response.write Serialize
	end sub
	
	
	' Helpers
	' Serializes a JSON object to JSON formatted string
	public function serializeObject(obj)
		dim out, prop, value, i, pairs, valueType
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
			
			if prop.name = JSON_ROOT_KEY then
				out = out & """" & obj.defaultPropertyName & """:"
			else
				out = out & """" & prop.name & """:"
			end if
			
			if isArray(value) or GetTypeName(value) = "JSONarray" then
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
	
	' Serializes a value to a valid JSON formatted string representing the value
	' (quoted for strings, the type name for objects, null for nothing and null values)
	public function serializeValue(byval value)
		dim out, offset

		select case lcase(GetTypeName(value))
			case "null"
				out = "null"
			
			case "nothing"
				out = "undefined"
			
			case "boolean"
				if value then
					out = "true"
				else
					out = "false"
				end if
			
			case "byte", "integer", "long", "single", "double", "currency", "decimal"
				out = value
			
			case "date"
				offset = GetTimeZoneOffset()
				
				out = """" & year(value) & "-" & padZero(month(value), 2) & "-" & padZero(day(value), 2) & "T" & padZero(hour(value), 2) & ":" & padZero(minute(value), 2) & ":" & padZero(second(value), 2) & left(offset, 1) & padZero(mid(offset, 2), 2) & ":00"""
			
			case "string", "char", "empty"
				out = """" & EscapeCharacters(value) & """"
			
			case else
				out = """" & GetTypeName(value) & """"
		end select
		
		serializeValue = out
	end function
	
	' Pads a numeric string with zeros at left
	private function padZero(byval number, byval length)
		dim result
		result = number
		
		while len(result) < length
			result = "0" & result
		wend
		
		padZero = result
	end function
	
	' Returns the time zone offset from the UTC time in hours (eg.: -3)
	private Function GetTimeZoneOffset()
		' http://ajaxandxml.blogspot.com.br/2006/02/computing-server-time-zone-difference.html
		%>
		<script runat="server" language="jscript">
			var JSON_TZDiff = new Date().getTimezoneOffset();
		</script>
		<%		
		GetTimeZoneOffset = - JSON_TZDiff / 60
	End Function
	
	' Serializes an array or JSONarray object to JSON formatted string
	public function serializeArray(byref arr)
		dim i, j, dimensions, out, innerArray, elm, val
		
		out = "["
		
		if isobject(arr) then
			innerArray = arr.items
		else
			innerArray = arr
		end if

		dimensions = NumDimensions(innerArray)
		
		for i = 1 to dimensions
			if i > 1 then out = out & ","
			
			if dimensions > 1 then out = out & "["
			
			for j = 0 to ubound(innerArray, i)
				if j > 0 then out = out & ","
				
				'multidimentional
				if dimensions > 1 then
					if isobject(innerArray(i - 1, j)) then
						set elm = innerArray(i - 1, j)
					else
						elm = innerArray(i - 1, j)
					end if
				else
					if isobject(innerArray(j)) then
						set elm = innerArray(j)
					else
						elm = innerArray(j)
					end if
				end if
								
				if isobject(elm) then
					if GetTypeName(elm) = "JSONobject" then
						set val = elm
					
					elseif GetTypeName(elm) = "JSONarray" then
						val = elm.items
						
					elseif isObject(elm.value) then
						set val = elm.value
						
					else
						val = elm.value
					end if
				else
					val = elm
				end if

				if isArray(val) or GetTypeName(val) = "JSONarray" then
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
	
	
	' Returns the number of dimensions an array has
	public Function NumDimensions(byref arr)
		Dim dimensions
		dimensions = 0
		
		On Error Resume Next
		
		Do While Err.number = 0
			dimensions = dimensions + 1
			UBound arr, dimensions
		Loop
		On Error Goto 0
		
		NumDimensions = dimensions - 1
	End Function
	
	' Pushes (adds) a value to an array, expanding it
	public function ArrayPush(byref arr, byref value)
		redim preserve arr(ubound(arr) + 1)
		
		if isobject(value) then
			set arr(ubound(arr)) = value
		else
			arr(ubound(arr)) = value
		end if
		
		ArrayPush = arr
	end function
	
	' Removes a value from an array
	private function ArraySlice(byref arr, byref index)
		dim i, upperBound
		i = index
		upperBound = ubound(arr)
		
		do while i < upperBound
			if isObject(arr(i)) then
				set arr(i) = arr(i + 1)
			else
				arr(i) = arr(i + 1)
			end if
			
			i = i + 1
		loop
		
		redim preserve arr(upperBound)
		
		ArraySlice = arr
	end function
	
	' Load properties from an ADO RecordSet object into an array
	' @param rs as ADODB.RecordSet
	public sub LoadRecordSet(byref rs)
		dim arr, obj, field
		
		set arr = new JSONArray
		
		while not rs.eof
			set obj = new JSONobject
		
			for each field in rs.fields
				obj.Add field.name, field.value
			next
			
			arr.Push obj
			
			rs.movenext
		wend
		
		set obj = nothing
		
		add JSON_ROOT_KEY, arr
	end sub
	
	' Load properties from the first record of an ADO RecordSet object
	' @param rs as ADODB.RecordSet
	public sub LoadFirstRecord(byref rs)
		dim field
		
		for each field in rs.fields
			add field.name, field.value
		next
	end sub
	
	' Returns the value's type name (usefull for types not supported by VBS)
	public function GetTypeName(byval value)
		dim valueType
	
		on error resume next
			valueType = TypeName(value)
			
			if err.number <> 0 then
				if varType(value) = 14 then valueType = "Decimal"
			end if
		on error goto 0
		
		GetTypeName = valueType
	end function
	
	' Escapes special characters in the text
	' @param text as String
	public function EscapeCharacters(byval text)
		dim result
		
		result = text
	
		if not isNull(text) then
			result = cstr(result)
			
			result = replace(result, "\", "\\")
			result = replace(result, """", "\""")
			result = replace(result, vbcr, "\r")
			result = replace(result, vblf, "\n")
			result = replace(result, vbtab, "\t")
		end if
	
		EscapeCharacters = result
	end function
	
	' Used to write the log messages to the response on debug mode
	private sub log(byval msg)
		if i_debug then response.write "<li>" & msg & "</li>" & vbcrlf
	end sub
end class


' JSON array class
' Represents an array of JSON objects and values
class JSONarray
	dim i_items, i_depth, i_parent, i_version, i_defaultPropertyName

	' The class version
	public property get version
		items = i_version
	end property

	' The actual array items
	public property get items
		items = i_items
	end property
	
	public property let items(value)
		if isArray(value) then
			i_items = value
		else
			err.raise JSON_ERROR_NOT_AN_ARRAY, TypeName(me), "The value assigned is not an array."
		end if
	end property
	
	' The length of the array
	public property get length
		length = ubound(i_items) + 1
	end property
	
	' The depth of the array in the chain (starting with 1)
	public property get depth
		depth = i_depth
	end property
	
	' The parent object or array
	public property get parent
		set parent = i_parent
	end property
	
	public property set parent(value)
		set i_parent = value
		i_depth = i_parent.depth + 1
		i_defaultPropertyName = i_parent.defaultPropertyName
	end property
	
	' Gets/sets the default property name generated when loading recordsets and arrays (default "data")
	public property get defaultPropertyName
		defaultPropertyName = i_defaultPropertyName
	end property

	public property let defaultPropertyName(value)
		i_defaultPropertyName = value
	end property

	
	
	' Constructor and destructor
	private sub class_initialize
		i_version = "2.3.5"
		i_defaultPropertyName = JSON_DEFAULT_PROPERTY_NAME
		redim i_items(-1)
		i_depth = 0
	end sub
	
	private sub class_terminate
		dim i, j, js, dimensions
		
		dimensions = 0
		
		On Error Resume Next
		
		Do While Err.number = 0
			dimensions = dimensions + 1
			UBound i_items, dimensions
		Loop
		
		On Error Goto 0
		
		dimensions = dimensions - 1
		
		for i = 1 to dimensions
			for j = 0 to ubound(i_items, i)
				if dimensions = 1 then
					set i_items(j) = nothing
				else
					set i_items(i - 1, j) = nothing
				end if
			next
		next
	end sub
	
	' Adds a value to the array
	public sub Push(byref value)
		dim js, instantiated
		
		if typeName(i_parent) = "JSONobject" then
			set js = i_parent
			i_defaultPropertyName = i_parent.defaultPropertyName
		else
			set js = new JSONobject
			js.defaultPropertyName = i_defaultPropertyName
			instantiated = true
		end if
		
		js.ArrayPush i_items, value
		
		if instantiated then set js = nothing
	end sub
	
	' Load properties from a ADO RecordSet object
	public sub LoadRecordSet(byref rs)
		dim obj, field
		
		while not rs.eof
			set obj = new JSONobject
		
			for each field in rs.fields
				obj.Add field.name, field.value
			next
			
			Push obj
			
			rs.movenext
		wend
		
		set obj = nothing
	end sub

	' Returns the item at the specified index
	' @param index as int - the desired item index
	public default function ItemAt(byval index)
		dim len
		len = me.length
		
		if len > 0 and index < len then
			if isObject(i_items(index)) then
				set ItemAt = i_items(index)
			else
				ItemAt = i_items(index)
			end if
		else
			err.raise JSON_ERROR_INDEX_OUT_OF_BOUNDS, TypeName(me), "Index out of bounds."
		end if
	end function
	
	' Serializes this JSONarray object in JSON formatted string value
	' (uses the JSONobject.SerializeArray method)
	public function Serialize()
		dim js, out, instantiated, actualLCID
		
		actualLCID = Response.LCID
		Response.LCID = 1033
		
		if not isEmpty(i_parent) then
			if TypeName(i_parent) = "JSONobject" then
				set js = i_parent
				i_defaultPropertyName = i_parent.defaultPropertyName
			end if
		end if
		
		if isEmpty(js) then
			set js = new JSONobject
			js.defaultPropertyName = i_defaultPropertyName
			instantiated = true
		end if
		
		out = js.SerializeArray(me)
		
		if instantiated then set js = nothing
		
		Response.LCID = actualLCID
		
		Serialize = out
	end function
	
	' Writes the serialized array to the response
	public function Write()
		Response.Write Serialize()
	end function
end class


' JSON pair class
' represents a name/value pair of a JSON object
class JSONpair
	dim i_name, i_value
	dim i_parent
	
	' The name or key of the pair
	public property get name
		name = i_name
	end property
	
	public property let name(val)
		i_name = val
	end property
	
	' The value of the pair
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
	
	' The parent object
	public property get parent
		set parent = i_parent
	end property
	
	public property set parent(val)
		set i_parent = val
	end property
	
	
	' Constructor and destructor
	private sub class_initialize
	end sub
	
	private sub class_terminate
		if isObject(value) then set value = nothing
	end sub
end class
%>