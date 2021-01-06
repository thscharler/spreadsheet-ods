# 0.5.2
- Add useability for XmlTag and TextTag.
- Implement a few standard TextTag Wrappers for common text elements.

# 0.5.1
- Split up the unwieldy PageLayout into PageStyle and MasterPage, and add
  an example how to use those bastards. 

# 0.5.0
- Major reorg of styles. Replaced Styles with separate CellStyle, ColStyle, 
  RowStyle etc.
- Create CellStyleRef, ColStyleRef, etc to be used when relating to styles.
  Should add a little bit of safety here.
- Introduced submodules for style: stylemap, tabstop, units

# 0.4.2

- Allow the ODS version to be specified. This adds support for 
  ODS 1.3. 
-- Default version set to 1.3. 

# 0.4.1 

- Refine usage of Style::cell(), cell_mut(), table(), table_mut(), col(), col_mut(), 
  row(), row_mut(). Assert the correct style-family for access to these Attributes.
  
- Bug when writing empty lines, used wrong row-style.

- No row/column styles are written if they are beyond the range of the maximum
  used cell. This is a desired behaviour. To make it easier there is a 
  Value-Conversion from '()' to an empty cell-value.

