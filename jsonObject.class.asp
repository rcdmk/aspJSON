<%
' JSON object class 2.2.3 - October, 10th - 2012
'
' Licence:
' The MIT License (MIT)
' Copyright (c) 2012 RCDMK - rcdmk[at]hotmail[dot]com
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
const JSON_SPECIAL_VALUES_REGEX = "^(?:(?:t(?:r(?:ue?)?)?)|(?:f(?:a(?:l(?:se?)?)?)?)|(?:n(?:u(?:ll?)?))|(?:u(?:n(?:d(?:e(?:f(?:i(?:n(?:ed?)?)?)?)?)?)?)?))$"

const JSON_ERROR_PARSE = 1
const JSON_ERROR_PROPERTY_ALREADY_EXISTS = 2
const JSON_ERROR_PROPERTY_DOES_NOT_EXISTS = 3
const JSON_ERROR_NOT_AN_ARRAY = 4

class JSONobject
	dim i_debug, i_depth, i_parent
	dim i_properties, i_version

	' Set to true to show the internals of the parsing mecanism
	public property get debug
		debug = i_debug
	end property
	
	public property let debug(value)
		i_debug = value
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
		i_version = "2.2.3"
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
	
	
	' Parse a JSON string and populate the object
	public function parse(byval strJson)
		dim regex, i, size, char, prevchar, quoted
		dim mode, item, key, value, openArray, openObject
		dim actualLCID, tmpArray, tmpObj, addedToArray
		dim root, currentObject, currentArray
		
		log("Load string: """ & strJson & """")
		
		' Store the actual LCID and use the en-US to conform with the JSON standard
		actualLCID = session.LCID
		session.LCID = 1033
		
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
				if key = JSON_ROOT_KEY and GetTypeName(currentArray) <> "JSONarray" then redim i_properties(-1)
				
				' Init object
				if char = "{" then
					log("Create object<ul>")
					
					if key <> JSON_ROOT_KEY or GetTypeName(root) = "JSONarray" then
						' creates a new object
						set item = new JSONobject
						set item.parent = currentObject
						
						addedToArray = false
						
						' Object is inside an array
						if GetTypeName(currentArray) = "JSONarray" then
							if currentArray.depth > currentObject.depth then
								' Add it to the array
								set item.parent = currentArray
								tmpArray = currentArray.items
								
								ArrayPush tmpArray, item
								
								currentArray.items = tmpArray
								
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
					if key = JSON_ROOT_KEY then set root = item
					
					addedToArray = false					
					
					' Array is inside an array
					if isobject(currentArray) and openArray > 0 then
						if currentArray.depth > currentObject.depth then
							' Add it to the array
							set item.parent = currentArray
							tmpArray = currentArray.items
							
							ArrayPush tmpArray, item
							
							currentArray.items = tmpArray
							
							addedToArray = true
							
							log("Added to parent array")
						end if
					end if
					
					if not addedToArray then
						set item.parent = currentObject
						currentObject.add key, item
						log("Added to parent object")
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
				
				' If it begins with a double quote, its a string value
				if char = """" then
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
					else
						value = value & char
					end if
				else
					' possible boolean and null values
					if regex.pattern <> JSON_SPECIAL_VALUES_REGEX then regex.pattern = JSON_SPECIAL_VALUES_REGEX
					if regex.test(char) or regex.test(value) then
						value = value & char
						if value = "true" or value = "false" or value = "null" or value = "undefined" then mode = "addValue"
					else
						' If is a numeric char
						if regex.pattern <> "\d" then regex.pattern = "\d"
						if regex.test(char) then
							value = value & char
						
						' If it's not a numeric char, but the prev char was a number
						' used to catch separators and special numeric chars
						elseif regex.test(prevchar) then
							if char = "." or lcase(char) = "e" or lcase(char) = "-" or lcase(char) = "+" then
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
						if GetTypeName(currentArray.parent) = "JSONarray" then
							set currentArray = currentArray.parent
						
						' if the parent is an object
						elseif GetTypeName(currentArray.parent) = "JSONobject" then
							set tmpObj = currentArray.parent
							
							' we search for the next parent array to set the current array
							while isObject(tmpObj) and GetTypeName(tmpObj) = "JSONobject"
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
						if GetTypeName(currentObject.parent) = "JSONobject" then
							set currentObject = currentObject.parent
						
						' If the parent is and array
						elseif GetTypeName(currentObject.parent) = "JSONarray" then
							set tmpObj = currentObject.parent
							
							' we search for the next parent object to set the current object
							while isObject(tmpObj) and GetTypeName(tmpObj) = "JSONarray"
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
		
		session.LCID = actualLCID
		
		set parse = root
	end function
	
	' Add a new property (key-value pair)
	public sub add(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if isObject(p) then
			err.raise JSON_ERROR_PROPERTY_ALREADY_EXISTS, TypeName(me), "A property already exists with the name: " & prop & "."
		else
			dim item
			set item = new JSONpair
			item.name = prop
			set item.parent = me
			
			if isArray(obj) then
				dim item2
				set item2 = new JSONarray
				item2.items = obj
				set item2.parent = me

				set item.value = item2
				
			elseif isObject(obj) then
				set item.value = obj
			else
				item.value = obj
			end if

			ArrayPush i_properties, item
		end if
	end sub
	
	' Return the value of a property by its key
	public default function value(byval prop)
		dim p
		getProperty prop, p
		
		if isObject(p) then
			if isObject(p.value) then
				set value = p.value
			else
				value = p.value
			end if
		else
			err.raise JSON_ERROR_PROPERTY_DOES_NOT_EXISTS, TypeName(me), "Property " & prop & " doesn't exists in this object."
		end if
	end function
	
	' Change the value of a property
	' Creates the property if it didn't exists
	public sub change(byval prop, byval obj)
		dim p
		getProperty prop, p
		
		if isObject(p) then
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
	
	' Returns a property if it exists
	' @param prop as string - the property name
	' @param out outProp as variant - will be filled with the property value, null if not found
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
	
	
	' Serialize the current object to a JSON formatted string
	public function Serialize()
		dim actualLCID, out
		actualLCID = session.LCID
		session.LCID = 1033
		
		out = serializeObject(me)
		
		session.LCID = actualLCID
		
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
				out = out & """data"":"
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
				out = """" & value & """"
			
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
	
	' By DavidRR: http://stackoverflow.com/a/13980554/1046610
	' 
	private Function GetTimeZoneOffset()
		Const sComputer = "."

		Dim oWmiService : Set oWmiService = _
			GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")

		Dim cItems : Set cItems = oWmiService.ExecQuery("SELECT * FROM Win32_ComputerSystem")

		Dim oItem
		
		For Each oItem In cItems
			GetTimeZoneOffset = oItem.CurrentTimeZone / 60
			Exit For
		Next
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
	private Function NumDimensions(byref arr) 
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
	
	' Load properties from a ADO RecordSet object
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
		
		add "data", arr
	end sub
	
	' returns the value's type name
	public function GetTypeName(byval value)
		dim valueType
	
		on error resume next
			valueType = typeName(value)
			
			if err.number <> 0 then
				if varType(value) = 14 then valueType = "Decimal"
			end if
		on error goto 0
		
		GetTypeName = valueType
	end function
	
	' Used to write the log messages to the response on debug mode
	private sub log(byval msg)
		if i_debug then response.write "<li>" & msg & "</li>" & vbcrlf
	end sub
end class


' JSON array class
' Represents an array of JSON objects and values
class JSONarray
	dim i_items, i_depth, i_parent, i_version

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
	end property
	
	
	' Constructor and destructor
	private sub class_initialize
		i_version = "2.2.1"
		redim i_items(-1)
		i_depth = 0
	end sub
	
	private sub class_terminate
		dim i
		for i = 0 to ubound(i_items)
			set i_items(i) = nothing
		next
	end sub
	
	' Adds a value to the array
	public sub Push(byref value)
		dim js, instantiated
		
		if not isEmpty(i_parent) then
			set js = i_parent
		else
			set js = new JSONobject
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
	
	' Serializes this JSONarray object in JSON formatted string value
	' (uses the JSONobject.SerializeArray method)
	public function Serialize()
		dim js, out, instantiated
		
		if not isEmpty(i_parent) then
			if TypeName(i_parent) = "JSONobject" then set js = i_parent
		end if
		
		if isEmpty(js) then
			set js = new JSONobject
			instantiated = true
		end if
		
		out = js.SerializeArray(me)
		
		if instantiated then set js = nothing
		
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