<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="dictionary.class.asp"-->
<%
Response.Write("--- Starting test --- <br>")

Response.Write("--- Initialize and test dictionary status --- <br>")

Dim dic

Set dic = new dictionary

Response.Write("--- Add values --- <br>")
dic.add_element "0001","Ice ice baby"
dic.add_element "0002","Dove lo vuoi tu"
dic.add_element "0003","Hot shot"
dic.add_element "0004","7 minuti in paradiso"
dic.add_element "0005","Dimmi dove e quando"
dic.add_element "0006","Balla per me"
dic.add_element "0007","Nudi e crudi"
dic.add_element "0008","Obbligo o verita'"
dic.add_element "0009","Sotto il vestito, niente"
dic.add_element "0010","Colazione a letto "
dic.add_element "0011","Luce dei miei occhi"
dic.add_element "0012","Comprami un giocattolo"
dic.add_element "0013","Facciamo pace"
dic.add_element "0014","Less is more"
dic.add_element "0015","call me baby"
dic.add_element "0016","Pensaci tu"
dic.add_element "0017","Souvenir d'amour"
dic.add_element "0018","Proca a prendermi"
dic.add_element "0019","Meet me at the hotel"
dic.add_element "0020","Sexy movie night"
dic.add_element "0021","Doccia bollente"
dic.add_element "0022","Blind date"
dic.add_element "0023","Oggi voglio..."
dic.add_element "0024","Nuovo blocchetto"
dic.write()

Response.Write("--- Check dictionary dimension --- <br>")
Response.write("Dictionary dimension: " & dic.get_dimension() & "<br>")

Response.Write("--- Check if the key 0010 has been used --- <br>")
Response.Write("Has 0010 been used? " & dic.check_if_key_has_been_used("0010") & "<br>")

Response.Write("--- Get dictionary dimension --- <br>")
Response.Write("Dimension: " & dic.get_dimension() & "<br>")

Response.Write("--- Get dictionary value from 4 position --- <br>")
Response.Write("Value: " & dic.get_value_from_index(4) & "<br>")

Response.Write("--- Get dictionary key from 4 position --- <br>")
Response.Write("Key: " & dic.get_key_from_index(4) & "<br>")

Response.Write("--- Get dictionary value of 0011 --- <br>")
Response.Write("Value: " & dic.get_value_from_key("0011") & "<br>")

Response.Write("--- Set dictionary value of 0011 --- <br>")
dic.set_value_from_key "0011","Sole della mia vita"
Response.Write("Value at 0011: " & dic.get_value_from_key("0011") & "<br>")

Response.Write("--- Change 0024 in 0034 --- <br>")
dic.change_key "0024","0034"
Response.Write("Value at 0034: " & dic.get_value_from_key("0034") & "<br>")

Response.Write("--- Remove dictionary 0034 element --- <br>")
dic.remove_element_from_key "0034"
dic.write()

Response.Write("--- Remove dictionary 22 element --- <br>")
dic.remove_element_from_index 22
dic.write()

Response.Write("--- Remove dictionary 13 element --- <br>")
dic.remove_element_from_index 13
dic.write()

Response.Write("--- Remove last element from dictionary --- <br>")
dic.remove_last_element
dic.write()

Response.Write("--- Get 0014 index from dictionary --- <br>")
dic.get_key_index "0015"
Response.Write("Index: " & dic.get_key_index ("0015") & "<br>")

Response.Write("--- Check if -Hot shot- value in the dictionary --- <br>")
Response.Write("Is present? " & dic.check_if_value_is_present ("Hot shot") & "<br>")

Response.Write("--- Check if -Hot shot- index the dictionary --- <br>")
Response.Write("Index: " & dic.get_first_value_index_occurrence ("Hot shot") & "<br>")

Response.Write("--- Manage dictionary for next text --- <br>")
Response.Write("- Change 0011 in -Hot shot- <br>")
dic.set_value_from_key "0011","Hot shot"
Response.Write("- Add 0022 with -Hot shot- <br>")
dic.add_element "0022","Hot shot"
dic.write()
Response.Write("- Retrieve all indices -Hot shot- <br>")
Response.Write("Indices: ")
Dim temp 
For Each temp in dic.get_all_value_indices("Hot shot")
Response.Write(" "&temp&" ")
Next
Response.Write("<br>")

Response.Write("--- Replace all -Hot shot- occurrences with -Hot kiss- in dictionary--- <br>")
dic.replace_all_value_occurrences "Hot shot","Hot kiss"
dic.write()

Response.Write("--- Remove all -Hot kiss- occurrences from dictionary--- <br>")
dic.remove_elements_from_value "Hot kiss"
dic.write()

Response.Write("--- Remove 10 15 5 element from dictionary--- <br>")
Dim temp_array(2)
temp_array(0) = 10
temp_array(1) = 15
temp_array(2) = 5
dic.remove_elements_from_indices temp_array
dic.write()
%>
