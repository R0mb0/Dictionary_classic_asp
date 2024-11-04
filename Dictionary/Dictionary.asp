<%
'------- Variable part--------'

'Fixed array'
Dim fixed_array(1)
fixed_array(0) = Null
fixed_array(1) = Null

'------- Public functions --------'

'Function to initialize the Dictionary'
Function get_initialized_dictionary()
Dim temp_array(0)
temp_array(0) = fixed_array
get_initialized_dictionary = temp_array
End Function

'Function to get the requested key value from dictionary using the index'
Function get_dictionary_key_from_index(dictionary,idx)
Dim temp
temp = UBound(dictionary)
If idx >=0 and idx <= temp Then
get_dictionary_key_from_index = dictionary(idx)(0)
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - get_dictionary_key_from_index", "Index error")
End If
End Function

'Function to get the requested value from dictionary using the index'
Function get_dictionary_value_from_index(dictionary,idx)
Dim temp
temp = UBound(dictionary)
If idx >=0 and idx <= temp Then
get_dictionary_value_from_index = dictionary(idx)(1)
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - get_dictionary_value_from_index", "Index error")
End If
End Function

'Function to get dictionary dimension'
Function get_dictionary_dimension(dictionary)
get_dictionary_dimension = UBound(dictionary)
End Function

'Funtion to check if a key has been used'
Function check_if_key_has_been_used(dictionary,key)
Dim temp
For Each temp In dictionary
If temp(0) = key Then 
check_if_key_has_been_used = true
Exit Function
End If
Next
check_if_key_has_been_used = false
End Function

'Function to add an element'
Function add_element_to_dynamic_array(dictionary,key,value)
Dim temp
temp = UBound(dictionary)
If temp = 0 and IsNull(dictionary(0)(0)) and IsNull(dictionary(0)(1)) Then 'If the dictionary has been just initializated'
dictionary(temp)(0) = key
dictionary(temp)(1) = value 
Else'If ther's no special case'
If Not check_if_key_has_been_used(dictionary,key) Then
temp = temp + 1
Redim Preserve dictionary(temp)
dictionary(temp) = fixed_array
dictionary(temp)(0) = key
dictionary(temp)(1) = value
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - add_element_to_dynamic_array", "Duplicated key")
End If 
End If
End Function

'Function to get value from key'
Function get_dictionary_value_from_key(dictionary,key)
If check_if_key_has_been_used(dictionary,key) Then
Dim temp
For Each temp In dictionary
If temp(0) = key Then
get_dictionary_value_from_key = temp(1)
Exit Function 
End If
Next
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - get_dictionary_value_from_key", "The key is not present")
End If
End Function

'Function to set value from key'
Function set_dictionary_value_from_key(dictionary,key,value)
If check_if_key_has_been_used(dictionary,key) Then
Dim temp
Dim temp_index
temp_index = 0
For Each temp In dictionary
If temp(0) = key Then
dictionary(temp_index)(1) = value
Exit Function 
End If
inedx_temp = temp_index + 1
Next
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - set_dictionary_value_from_key", "The key is not present")
End If
End Function

'Function to change dictionary key'
Function change_dictionary_key(dictionary,old_key,new_key)
If check_if_key_has_been_used(dictionary,old_key) Then
If Not check_if_key_has_been_used(dictionary,new_key) Then 
Dim temp 
Dim temp_index
temp_index = 0
For Each temp In dictionary
If temp(0) = old_key Then 
dictionary(temp_index)(0) = new_key
End If
temp_index = temp_index + 1
Next
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - change_dictionary_key", "The new key is used")
End If
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - change_dictionary_key", "The old key is not present")
End If
End Function

'Function to remove a dictionary item from key'
Function remove_dictionary_element_from_key(dictionary,key)
If check_if_key_has_been_used(dictionary,key) Then 
Dim temp_array
temp_array = Array()
Dim temp_index
temp_index = 0
Dim temp 
For Each temp In dictionary
If Not temp(0) = key Then 
Redim  Preserve temp_array(temp_index)
temp_array(temp_index) = temp
End If
temp_index = temp_index + 1
Next
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - remove_dictionary_element_from_key", "The key is not present")
End If 
dictionary = temp_array
End Function

'Function to remove a dictionary item from key'
Function remove_dictionary_element_from_index(dictionary,idx)
Dim temp 
temp = UBound(dictionary)
If idx >=0 and idx <= temp Then 
Dim temp_index
temp_index = 0
Dim temp_array
temp_array = Array()
Dim temp_item
For Each temp_item In dictionary
If Not temp_inedx = idx Then 
Redim Preserve temp_array(temp_index)
temp_array(temp_index) = temp_item
End If
temp_index = temp_index + 1 
Next
Else
Call Err.Raise(vbObjectError + 10, "Dictionary - remove_dictionary_element_from_index", "Index error")
End If
dictionary = temp_array
End Function

'------Last Functions------'

'Funtion to write the entire dictionary'
Function write_dictionary(dictionary)
Dim temp
Dim temp_index
temp_index = 0
Response.Write("------Dictionary items------ <br><br>")
For Each temp In dictionary
Response.Write("Index: " & temp_index & "<br>")
Response.Write("Key: " & temp(0) & "<br>")
Response.Write("Value: " & temp(1) & "<br>")
Response.Write("------ <br>")
temp_index = temp_index + 1
Next
End Function
%>