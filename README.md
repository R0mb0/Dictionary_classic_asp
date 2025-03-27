# Dynamic array in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/5fb4b710ca5c4dfd88a883944af2dac3)](https://app.codacy.com/gh/R0mb0/Dynamic_array_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Dynamic_array_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Dynamic_array_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)


<details>
  <summary> 

## `Dynamic_array.asp`'s avaible Functions
  
</summary>

- **Initialize the array** -> `initialize_array()`
- **Check if the array is initializated** -> `is_array_initializated()`
- **Add an element into the array** -> `add_element_to_array(element)`
- **Get element from index** -> `get_element_from_array(idx)`
- **Remove last element from array** -> `remove_last_element_from_array()`
- **Remove an element from array** -> `remove_all_occurences_from_array(element)`
- **Remove the first element occurence from array** -> `remove_first_occurence_from_array(element)`
- **Remove element from index** -> `remove_this_elements_from_array(idx)`
- **Remove elements from indices** -> `remove_these_elements_from_array(indices_array)`
- **Reset the array** -> `initialize_array()`
- **Check if an element is in the array** -> `array_contains(element)`
- **Retrieve the first index of an element in the array** -> `from_array_get_first_index_occurence_of(element)`
- **Retrieve all indeces of an element inside the array (return an array)** -> `from_array_get_all_indeces_occurence_of(element)`
- **Retrieve the entire array** -> `get_array()`
- **Retrieve the array dimension** -> `get_array_dimension()`
- **Write the entire array** -> `write_array()`

## How to use 

> From `Test1.asp`

1. Initialize the array and check it's status
   ```asp
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="Dynamic_array.asp"-->
   <%
   initialize_array()
   Response.Write("Array status: ")
   Response.Write(is_array_initializated() & "<br>")
   ```
2. Use the functions to manage the array
   ```asp
   add_element_to_array("A")
   add_element_to_array("B")
   add_element_to_array("C")
   add_element_to_array("D")
   Response.Write("Elements inside: ")
   write_array()
   %>
   ```

</details>

<details>
  <summary> 

## `Dynamic_arrays.asp`'s avaible Functions
  
</summary>

- **Initialize a dynamic array** -> `get_initializated_dynamic_array()`
- **Add an element into a dynamic array** -> `add_element_to_dynamic_array(my_array,element)`
- **Get element from index** -> `get_element_from_dynamic_array(my_array,idx)`
- **Remove last element from a dynamic array** -> `remove_last_element_from_dynamic_array(my_array)`
- **Remove all element occurences from a dyamic array** -> `remove_all_occurences_from_dynamic_array(my_array,element)`
- **Remove first element occurence from a dynamic array** -> `remove_first_occurence_from_dynamic_array(my_array,element)`
- **Remove element from index** -> `remove_this_elements_from_dynamic_array(my_array,idx)`
- **Remove elements from indices** -> `remove_these_elements_from_dynamic_array(my_array,indices_array)`
- **Reset a dynamic array** -> `get_initializated_dynamic_array()`
- **Check if an element is in the dynamic array** -> `dynamic_array_contains(my_array,element)`
- **Retrieve the first index of an element inside a dynamic array** -> `from_dynamic_array_get_first_index_occurence_of(my_array,element)`
- **Retrieve all indeces of an element inside the dynamic array (return an array)** -> `from_dynamic_array_get_all_indeces_occurence_of(my_array,element)`
- **Retrieve the dynamic array dimension** -> `get_dynamic_array_dimension(my_array)`
- **Write an entire dynamic array** -> `write_dynamic_array(my_array)`

## How to use 

> From `Test.asp`

1. Create array and initialize it
   ```asp
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="Dynamic_arrays.asp"-->
   <%
   Dim test_array
   test_array = Array()
   test_array = get_initializated_dynamic_array()
   ```
2. Pass the array to functions for manage
   ```asp
   add_element_to_dynamic_array test_array,"A"
   add_element_to_dynamic_array test_array,"B"
   add_element_to_dynamic_array test_array,"C"
   write_dynamic_array(test_array)
   %>
   ```

</details>

<details>
  <summary> 

## `dArray.class.asp`'s avaible Functions
  
</summary>

- **Initialize class** -> `sub class_initialize()`
- **Terminate class** -> `sub class_terminate()`
- **Add an element** -> `Public Function add_element(element)`
- **Get element from index** -> `Public Function get_element(idx)`
- **Get array dimension** -> `Public Function get_dimension()`
- **Remove last element** -> `Public Function remove_last_element()`
- **Remove all element occurences** -> `Public Function remove_all_occurences(element)`
- **Remove first element occurence** -> `Public Function remove_first_occurence(element)`
- **Remove element from index** -> `Public Function remove_this_element(idx)`
- **Remove elements from indices** -> `Public Function remove_these_elements(indices_array)`
- **Reset** -> Re-initialize the class
- **Check if an element is present** -> `Public Function contains(element)`
- **Retrieve the first index of an element** -> `Public Function get_first_index_occurence_of(element)`
- **Retrieve all indeces of an element (return an array)** -> `Public Function get_all_indeces_occurence_of(element)`
- **Write the entire array** -> `Public Function write_array()`

## How to use 

> From `Test.asp`

1. Initialize the class
  ```asp
  <%@LANGUAGE="VBSCRIPT"%>
  <!--#include file="dArray.class.asp"-->
  <%
      Dim da 
      Set da = new dArray
  ```
2. Use the class
  ```asp
  da.add_element("A")
  da.add_element("B")
  da.add_element("C")
  da.add_element("D")
  da.add_element("A")
  da.add_element("B")
  da.add_element("C")
  da.add_element("D")
  da.add_element("A")
  Response.Write("Elements inside: ")
  da.write_array()
  %>
  ```

</details>
