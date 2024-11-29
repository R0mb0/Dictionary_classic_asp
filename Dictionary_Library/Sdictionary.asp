<%
'------- Variable part--------'

'Fixed array'
Dim fixed_array(1)
fixed_array(0) = Null
fixed_array(1) = Null
Dim my_dictionary
my_dictionary = Array()

Dim last_index_searched

'------- Public functions --------'

'Function to initialize the Dictionary'
Function get_initialized_dictionary()
  Dim temp_array(0)
  temp_array(0) = fixed_array
  my_dictionary = temp_array

End Function

'Function to get the requested key value from dictionary using the index'
Function get_dictionary_key_from_index(ByVal idx)
  Dim temp
  temp = UBound(my_dictionary)
  If idx >=0 and idx <= temp Then
    get_dictionary_key_from_index = my_dictionary(idx)(0)
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_dictionary_key_from_index", "Index error: "&idx&"")
  End If
End Function

'Function to get the requested value from dictionary using the index'
Function get_dictionary_value_from_index(ByVal idx)
  Dim temp
  temp = UBound(my_dictionary)
  If idx >=0 and idx <= temp Then
    get_dictionary_value_from_index = my_dictionary(idx)(1)
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_dictionary_value_from_index", "Index error: "&idx&"")
  End If
End Function

'Function to get dictionary dimension'
Function get_dictionary_dimension()
  get_dictionary_dimension = UBound(my_dictionary)
End Function

'Funtion to check if a key has been used'
Function check_if_key_has_been_used(ByVal key)
  Dim temp
  Dim temp_index
  temp_index = 0
  For Each temp In my_dictionary
    If temp(0) = key Then 
      check_if_key_has_been_used = true
      last_index_searched = temp_index
      Exit Function
    End If
    temp_index = temp_index + 1
  Next
  check_if_key_has_been_used = false
End Function

'Function to add an element'
Function add_element_to_dictionary_array(ByVal key,ByVal value)
  Dim temp
  temp = UBound(my_dictionary)
  If temp = 0 and IsNull(my_dictionary(0)(0)) and IsNull(my_dictionary(0)(1)) Then 'If the dictionary has been just initializated'
    my_dictionary(temp)(0) = key
    my_dictionary(temp)(1) = value 
  Else'If ther's no special case'
    If Not check_if_key_has_been_used(key) Then
      temp = temp + 1
      Redim Preserve my_dictionary(temp)
      my_dictionary(temp) = fixed_array
      my_dictionary(temp)(0) = key
      my_dictionary(temp)(1) = value
    Else
      Call Err.Raise(vbObjectError + 10, "Sdictionary - add_element_to_dictionary_array", "Duplicated key: "&key&"")
    End If 
  End If
End Function

'Function to get value from key'
Function get_dictionary_value_from_key(ByVal key)
  If check_if_key_has_been_used(key) Then
    get_dictionary_value_from_key = my_dictionary(last_index_searched)(1)
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_dictionary_value_from_key", "The key "&key&" is not present")
  End If
End Function

'Function to set value from key'
Function set_dictionary_value_from_key(ByVal key,ByVal value)
  If check_if_key_has_been_used(key) Then
    my_dictionary(last_index_searched)(1) = value
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - set_dictionary_value_from_key", "The key "&key&" is not present")
  End If
End Function

'Function to change dictionary key'
Function change_dictionary_key(ByVal old_key,ByVal new_key)
  If check_if_key_has_been_used(old_key) Then
    Dim temp
    temp = last_index_searched
    If Not check_if_key_has_been_used(new_key) Then 
      my_dictionary(temp)(0) = new_key
    Else
      Call Err.Raise(vbObjectError + 10, "Sdictionary - change_dictionary_key", "The new key "&new_key&" is used")
    End If
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - change_dictionary_key", "The old key "&old_key&" is not present")
  End If
End Function

'Function to remove a dictionary item from key'
Function remove_dictionary_element_from_key(ByVal key)
  If check_if_key_has_been_used(key) Then 
    Dim temp_array
    temp_array = Array()
    Dim temp_index
    temp_index = 0
    Dim temp 
      For Each temp In my_dictionary
        If Not temp(0) = key Then 
          Redim  Preserve temp_array(temp_index)
          temp_array(temp_index) = temp
        End If
        temp_index = temp_index + 1
      Next
    Else
      Call Err.Raise(vbObjectError + 10, "Sdictionary - remove_dictionary_element_from_key", "The key "&key&" is not present")
    End If 
      my_dictionary = temp_array
End Function

'Function to remove a dictionary item from index'
Function remove_dictionary_element_from_index(ByVal idx)
  Dim temp 
  temp = UBound(my_dictionary)
  If idx >=0 and idx <= temp Then 
    Dim temp_index
    temp_index = 0
    Dim temp_array_index
    temp_array_index = 0
    Dim temp_array
    temp_array = Array()
    Dim temp_item
    For Each temp_item In my_dictionary
      If Not temp_index = idx Then 
        Redim Preserve temp_array(temp_array_index)
        temp_array(temp_array_index) = temp_item
        temp_array_index = temp_array_index + 1
      End If
      temp_index = temp_index + 1 
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - remove_dictionary_element_from_index", "Index error: "&idx&"")
  End If
  my_dictionary = temp_array
End Function

'Funtion to remove last element from the dictionary'
Function remove_last_element_from_dictionary()
  Dim temp 
  temp = UBound(my_dictionary)
  If temp > 0 Then
    temp = temp - 1
    Redim Preserve my_dictionary(temp)
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - remove_last_element_from_dictionary", "No element to remove")
  End If
End Function

'Function to get key index from dictionary'
Function get_key_index_from_dictionary(ByVal key)
  If check_if_key_has_been_used(key) Then
    get_key_index_from_dictionary = last_index_searched
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_key_index_from_dictionary", "The key "&key&" is not present")
  End If 
End Function

'Function to check if a value is in the dictionary'
Function check_if_value_is_present(ByVal value)
  Dim temp
  temp_index = 0
  For Each temp In my_dictionary
    If temp(1) = value Then 
      check_if_value_is_present = true
      last_index_searched = temp_index
      Exit Function
    End If
    temp_index = temp_index + 1
  Next
  check_if_value_is_present = false
End Function

'Function to get first index value'
Function get_first_value_index_occurrence(ByVal value)
  If check_if_value_is_present(value) Then
    get_first_value_index_occurrence = last_index_searched
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_first_value_index_occurrence", "The value "&value&" is not present")
  End If
End Function

'Function to retrieve all value indices'
Function get_all_value_indices(ByVal value)
  If check_if_value_is_present(value) then
    Dim temp_array()
    Dim temp 
    Dim temp_index
    temp_index = 0
    Dim temp_array_index
    temp_array_index = 0
    For Each temp In my_dictionary
      If temp(1) = value Then 
        Redim Preserve temp_array(temp_array_index)
        temp_array(temp_array_index) = temp_index
        temp_array_index = temp_array_index + 1
      End If
      temp_index = temp_index + 1
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - get_first_value_index_occurrence", "The value "&value&" is not present")
End If
  get_all_value_indices = temp_array
End Function

'Function to remove elements from an array of indices (pass an array with indices)'
Function remove_dictionary_elements_from_indices(ByVal indices)
  Dim dimension
  dimension = UBound(my_dictionary)
  Dim temp_array
  temp_array = Array()
  Dim temp 
  For Each temp In indices
    If temp >= 0 and temp <= dimension Then 
      my_dictionary(temp) = Null
    Else 
      Call Err.Raise(vbObjectError + 10, "Sdictionary - remove_dictionary_elements_from_indices", "Index error: "&temp&"")
    End If
  Next
  Dim temp_index
  temp_index = 0
  For Each temp In my_dictionary
    If Not IsNull(temp) Then 
      Redim Preserve temp_array(temp_index)
      temp_array(temp_index) = temp
      temp_index = temp_index + 1
    End If
  Next
  my_dictionary = temp_array
End Function

'Function to remove all elements with that value (remove also one element if the value is unique)'
Function remove_dictionary_elements_from_value(ByVal value)
  Dim temp_array
  temp_array = Array()
  temp_array = get_all_value_indices(value)
  Dim temp 
  For Each temp In temp_array
    my_dictionary(temp) = Null
  Next
  Dim temp_index
  temp_index = 0
  For Each temp In my_dictionary
    If Not IsNull(temp) Then 
      Redim Preserve temp_array(temp_index)
      temp_array(temp_index) = temp
      temp_index = temp_index + 1
    End If
  Next
  my_dictionary = temp_array
End Function

'Function to replace all value occurrences with new value'
Function replace_all_value_occurrences(ByVal old_value,ByVal new_value)
  If check_if_value_is_present(old_value) Then 
    Dim temp 
    Dim temp_index
    temp_index = 0
    For Each temp In my_dictionary
      If temp(1) = old_value Then 
        my_dictionary(temp_index)(1) = new_value
      End If
      temp_index = temp_index + 1
    Next
  Else
    Call Err.Raise(vbObjectError + 10, "Sdictionary - replace_all_value_occurrences", "The old value "&old_value&" is not present")
  End If
End Function

'Funtion to write the entire dictionary'
Function write_dictionary()
  Dim temp
  Dim temp_index
  temp_index = 0
  Response.Write("------Dictionary items------ <br><br>")
  For Each temp In my_dictionary
    Response.Write("Index: " & temp_index & "<br>")
    Response.Write("Key: " & temp(0) & "<br>")
    Response.Write("Value: " & temp(1) & "<br>")
    Response.Write("------ <br>")
  temp_index = temp_index + 1
  Next
End Function
%>
