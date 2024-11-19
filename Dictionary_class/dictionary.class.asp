<%

Class dictionary

Dim fixed_array()
Dim last_index_searched
Dim my_dictionary

' Initialization and destruction'
	sub class_initialize()
        Redim fixed_array(1)
        fixed_array(0) = Null
        fixed_array(1) = Null
        Dim temp_array(0)
        temp_array(0) = fixed_array
        my_dictionary = temp_array
	end sub
	
	sub class_terminate()
		' Destroy all elements of the arrays'
		redim lines(-1)
		redim headers(-1)
	end sub

end Class

%>