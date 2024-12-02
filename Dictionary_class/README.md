# Dictionaries class in Classic ASP

## `dictionary.class.asp`'s avaible Functions

- **Add an element into a dictionary** -> `add_element(key,value)`
- **Get value from key** -> `get_value_from_key(key)`
- **Set value from key** -> `set_value_from_key(key,value)`
- **Change a key into dictionary** -> `change_key(old_key,new_key)`
- **Remove dictionary element from key** -> `remove_element_from_key(key)`
- **Remove dictionary element from index** -> `remove_element_from_index(idx)`
- **Remove last element from dictionary** -> `remove_last_element(dictionary)`
- **Remove elements from dictionary using indices (pass to function an array with the indices)** -> `remove_elements_from_indices(indices)`
- **Remove all elements with the same value from dictionary** -> `remove_elements_from_value(value)`
- **Replace all value occurrences with a new value in dictionary** -> `replace_all_value_occurrences(old_value,new_value)`
- **Get key index from dictionary** -> `get_key_index(key)`
- **Reset a dictionary** -> Re-initialize the object
- **Get dictionary key from index** -> `get_key_from_index(idx)`
- **Get first value index occurrence** -> `get_first_value_index_occurrence(value)`
- **Get all indices of a value** -> `get_all_value_indices(value)`
- **Check if a key has been used into dictionary** -> `check_if_key_has_been_used(key)`
- **Check if a value is present in the dictionary** -> `check_if_value_is_present(value)`
- **Get dictionary value from index** -> `get_value_from_index(idx)`
- **Get dictionary dimension** -> `get_dimension()`
- **Write a dictionary** -> `write()`

## How to use: 

> From `Test.asp`

1. Initialize the class
  ```
  <%@LANGUAGE="VBSCRIPT"%>
  <!--#include file="dictionary.class.asp"-->
  <%
      Dim dic
      Set dic = new dictionary
  ```
2. Use the class
  ```
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
  %>
  ```
