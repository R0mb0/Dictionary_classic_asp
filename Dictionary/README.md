# Backgroud structure

## Arrays structure.
The dictionary is implemented as a multidimensional array, where the first dynamic array contains on every cell a fixed array of two cells. 
The first fixed array cell contains the unique key and the second one contains the value. 

### Chart
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
