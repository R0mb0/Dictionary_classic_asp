<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Dictionary.asp"-->
<%
Response.Write("--- Starting test --- <br>")

Response.Write("--- Initialize dictionary --- <br>")

Dim dictionary
dictionary = Array()

dictionary = get_initialized_dictionary()

Response.Write("--- Add values --- <br>")
add_element_to_dynamic_array dictionary,"0001","Ice ice baby"
add_element_to_dynamic_array dictionary,"0002","Dove lo vuoi tu"
add_element_to_dynamic_array dictionary,"0003","Hot shot"
add_element_to_dynamic_array dictionary,"0004","7 minuti in paradiso"
add_element_to_dynamic_array dictionary,"0005","Dimmi dove e quando"
add_element_to_dynamic_array dictionary,"0006","Balla per me"
add_element_to_dynamic_array dictionary,"0007","Nudi e crudi"
add_element_to_dynamic_array dictionary,"0008","Obbligo o verita'"
add_element_to_dynamic_array dictionary,"0009","Sotto il vestito, niente"
add_element_to_dynamic_array dictionary,"0010","Colazione a letto "
add_element_to_dynamic_array dictionary,"0011","Luce dei miei occhi"
add_element_to_dynamic_array dictionary,"0012","Comprami un giocattolo"
add_element_to_dynamic_array dictionary,"0013","Facciamo pace"
add_element_to_dynamic_array dictionary,"0014","Less is more"
add_element_to_dynamic_array dictionary,"0015","call me baby"
add_element_to_dynamic_array dictionary,"0016","Pensaci tu"
add_element_to_dynamic_array dictionary,"0017","Souvenir d'amour"
add_element_to_dynamic_array dictionary,"0018","Proca a prendermi"
add_element_to_dynamic_array dictionary,"0019","Meet me at the hotel"
add_element_to_dynamic_array dictionary,"0020","Sexy movie night"
add_element_to_dynamic_array dictionary,"0021","Doccia bollente"
add_element_to_dynamic_array dictionary,"0022","Blind date"
add_element_to_dynamic_array dictionary,"0023","Oggi voglio..."
add_element_to_dynamic_array dictionary,"0024","Nuovo blocchetto"
write_dictionary dictionary

Response.Write("--- Check if the key 0010 has been used --- <br>")
Response.Write("Has 0010 been used? " & check_if_key_has_been_used(dictionary,"0010") & "<br>")

Response.Write("--- Get dictionary dimension --- <br>")
Response.Write("Dimension: " & get_dictionary_dimension(dictionary) & "<br>")

Response.Write("--- Get dictionary value from 4 position --- <br>")
Response.Write("Value: " & get_dictionary_value_from_index(dictionary,4) & "<br>")

Response.Write("--- Get dictionary key from 4 position --- <br>")
Response.Write("Key: " & get_dictionary_key_from_index(dictionary,4) & "<br>")

Response.Write("--- Get dictionary value of 0011 --- <br>")
Response.Write("Value: " & get_dictionary_value_from_key(dictionary,"0011") & "<br>")

Response.Write("--- Set dictionary value of 0011 --- <br>")
set_dictionary_value_from_key dictionary,"0011","Sole della mia vita"
Response.Write("Value at 0011: " & get_dictionary_value_from_key(dictionary,"0011") & "<br>")

Response.Write("--- Change 0024 in 0034 --- <br>")
change_dictionary_key dictionary,"0024","0034"
Response.Write("Value at 0034: " & get_dictionary_value_from_key(dictionary,"0034") & "<br>")

Response.Write("--- Remove dictionary 0034 element --- <br>")
remove_dictionary_element_from_key dictionary,"0034"
write_dictionary dictionary

Response.Write("--- Remove dictionary 22 element--- <br>")
remove_dictionary_element_from_index dictionary,22
write_dictionary dictionary
%>