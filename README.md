# Dictionary in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/aadc6762a5a64808a2887afea2969e1d)](https://app.codacy.com/gh/R0mb0/Dictionary_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Dictionary_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Dictionary_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## Backgroud structure

### Arrays structure.
The dictionary is implemented as a multidimensional array, where the first dynamic array contains on every cell a fixed array of two cells. 
The first fixed array cell contains the unique key and the second one contains the value. 

#### Chart
```mermaid
%%{
  init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#BB2528',
      'primaryTextColor': '#fff',
      'primaryBorderColor': '#7C0000',
      'lineColor': '#F8B229',
      'secondaryColor': '#006100',
      'tertiaryColor': '#fff'
    }
  }
}%%
graph TD;
    A[/0/] ==> B[/1/]
    B ==> C[/2/]
    C ==> D[/.../]
    A --> E[0: Unique Key]
    E --> F[1: Value]
    B --> G[0: Unique Key]
    G --> H[1: Value]
    C --> I[0: Unique Key]
    I --> J[1: Value]
    D --> K[0: Unique Key]
    K --> L[1: Value]
```

<details>
  <summary> 

 ## `Dictionary.asp`'s avaible Functions
  
  </summary>

- **Initialize an array as a dictionary** -> `get_initialized_dictionary()`
- **Add an element into a dictionary** -> `add_element_to_dictionary(key,value)`
- **Get value from key** -> `get_dictionary_value_from_key(key)`
- **Set value from key** -> `set_dictionary_value_from_key(key,value)`
- **Change a key into dictionary** -> `change_dictionary_key(old_key,new_key)`
- **Remove dictionary element from key** -> `remove_dictionary_element_from_index(idx)`
- **Remove dictionary element from index** -> `remove_dictionary_element_from_index(idx)`
- **Remove last element from dictionary** -> `remove_last_element_from_dictionary()`
- **Remove elements from dictionary using indices (pass to function an array with the indices)** -> `remove_dictionary_elements_from_indices(indices)`
- **Remove all elements with the same value from dictionary** -> `remove_dictionary_elements_from_value(value)`
- **Replace all value occurrences with a new value in dictionary** -> `replace_all_value_occurrences(old_value,new_value)`
- **Get key index from dictionary** -> `get_key_index_from_dictionary(key)`
- **Reset a dictionary** -> `get_initialized_dictionary()`
- **Get dictionary key from index** -> `get_dictionary_key_from_index(idx)`
- **Get first value index occurrence** -> `get_first_value_index_occurrence(value)`
- **Get all indices of a value** -> `get_all_value_indices(value)`
- **Check if a key has been used into dictionary** -> `check_if_key_has_been_used(key)`
- **Check if a value is present in the dictionary** -> `check_if_value_is_present(value)`
- **Get dictionary value from index** -> `get_dictionary_value_from_index(idx)`
- **Get dictionary dimension** -> `get_dictionary_dimension()`
- **Write a dictionary** -> `write_dictionary()`

## How to use: 

> From `Test.asp`

1. Create array and initialize it as dictionary
   ```
   <!--#include file="Dictionaries.asp"-->
   <%
   Dim dictionary
   dictionary = Array()
   dictionary = get_initialized_dictionary()
   ```
2. Pass the dictionary to functions for manage
   ```
   add_element_to_dictionary_array "0001","Ice ice baby"
   add_element_to_dictionary_array "0002","Dove lo vuoi tu"
   add_element_to_dictionary_array "0003","Hot shot"
   write_dictionary
   %>
   ```
</details>

<details>
  <summary> 

 ## `Dictionaries.asp`'s avaible Functions
  
  </summary>

- **Initialize an array as a dictionary** -> `get_initialized_dictionary()`
- **Add an element into a dictionary** -> `add_element_to_dictionary(dictionary,key,value)`
- **Get value from key** -> `get_dictionary_value_from_key(dictionary,key)`
- **Set value from key** -> `set_dictionary_value_from_key(dictionary,key,value)`
- **Change a key into dictionary** -> `change_dictionary_key(dictionary,old_key,new_key)`
- **Remove dictionary element from key** -> `remove_dictionary_element_from_index(dictionary,idx)`
- **Remove dictionary element from index** -> `remove_dictionary_element_from_index(dictionary,idx)`
- **Remove last element from dictionary** -> `remove_last_element_from_dictionary(dictionary)`
- **Remove elements from dictionary using indices (pass to function an array with the indices)** -> `remove_dictionary_elements_from_indices(dictionary,indices)`
- **Remove all elements with the same value from dictionary** -> `remove_dictionary_elements_from_value(dictionary,value)`
- **Replace all value occurrences with a new value in dictionary** -> `replace_all_value_occurrences(dictionary,old_value,new_value)`
- **Get key index from dictionary** -> `get_key_index_from_dictionary(dictionary,key)`
- **Reset a dictionary** -> `get_initialized_dictionary()`
- **Get dictionary key from index** -> `get_dictionary_key_from_index(dictionary,idx)`
- **Get first value index occurrence** -> `get_first_value_index_occurrence(dictionary,value)`
- **Get all indices of a value** -> `get_all_value_indices(dictionary,value)`
- **Check if a key has been used into dictionary** -> `check_if_key_has_been_used(dictionary,key)`
- **Check if a value is present in the dictionary** -> `check_if_value_is_present(dictionary,value)`
- **Get dictionary value from index** -> `get_dictionary_value_from_index(dictionary,idx)`
- **Get dictionary dimension** -> `get_dictionary_dimension(dictionary)`
- **Write a dictionary** -> `write_dictionary(dictionary)`

## How to use: 

> From `Test.asp`

1. Create array and initialize it as dictionary
   ```
   <!--#include file="Dictionary.asp"-->
   <%
   Dim dictionary
   dictionary = Array()
   dictionary = get_initialized_dictionary()
   ```
2. Pass the dictionary to functions for manage
   ```
   add_element_to_dictionary_array dictionary,"0001","Ice ice baby"
   add_element_to_dictionary_array dictionary,"0002","Dove lo vuoi tu"
   add_element_to_dictionary_array dictionary,"0003","Hot shot"
   write_dictionary dictionary
   %>
   ```
   </details>

<details>
  <summary> 

## `dictionary.class.asp`'s avaible Functions
  
  </summary>

- **Add an element into a dictionary** -> `add_element(key,value)`
- **Get value from key** -> `get_value_from_key(key)`
- **Set value from key** -> `set_value_from_key(key,value)`
- **Change a key into dictionary** -> `change_key(old_key,new_key)`
- **Remove dictionary element from key** -> `remove_element_from_index(idx)`
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
</details>
