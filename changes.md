
# 0.4.1 

- Refine usage of Style::cell(), cell_mut(), table(), table_mut(), col(), col_mut(), 
  row(), row_mut(). Assert the correct style-family for access to these Attributes.
  
- Bug when writing empty lines, used wrong row-style.

- No row/column styles are written if they are beyond the range of the maximum
  used cell. This is a desired behaviour. To make it easier there is a 
  Value-Conversion from '()' to an empty cell-value.

