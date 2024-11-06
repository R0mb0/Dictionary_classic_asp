# Dictionary in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/aadc6762a5a64808a2887afea2969e1d)](https://app.codacy.com/gh/R0mb0/Dictionary_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Dictionary_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Dictionary_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `Dictionary.asp`'s avaible Functions

- **Initialize an array as a dictionary** -> `get_initialized_dictionary()`
- **Add an element into a dictionary** -> `add_element_to_dictionary_array(dictionary,key,value)`
- **Get value from key** -> `get_dictionary_value_from_key(dictionary,key)`
- **Set value from key** -> `set_dictionary_value_from_key(dictionary,key,value)`
- **Change a key into dictionary** -> `change_dictionary_key(dictionary,old_key,new_key)`
- **Remove dictionary element from key** -> `remove_dictionary_element_from_index(dictionary,idx)`
- **Remove dictionary element from index** -> `remove_dictionary_element_from_index(dictionary,idx)`
- **Remove last element from dictionary** -> `remove_last_element_from_dictionary(dictionary)`
- **Get key index from dictionary** -> `get_key_index_from_dictionary(dictionary,key)`
- **Reset a dictionary** -> `get_initialized_dictionary()`
- **Get dictionary key from index** -> `get_dictionary_key_from_index(dictionary,idx)`
- **Check if a key has been used into dictionary** -> `check_if_key_has_been_used(dictionary,key)`
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
