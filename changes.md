# 0.8.2

- Checks that formulas start with "of:="
- New f*ref variants for formulas. These create a diverse array of absolute $
  references.
- Value can extract a NaiveDate value.

WorkBook:

- Add sheet_idx to find a sheet by name.
- Add used_cols, used_rows to find the number of row/column-headers.
- fixed a missing namespace.

Sheet:

- clear_formula, clear_cell_style, clear_validation

# 0.8.1

- fix for #24. Excel doesn't like the empty content-validations tag.
- complex text was missing the value-type
- Value conversions from ref types.

# 0.8.0

- Value::TextXml changed to Vec<TextTag>. Multiline, styled text can occur as
  multiple direct text:p in a table:cell.

- Repeat rows where not correctly considered when occuring in the middle of the
  table.

  There is a problem that still remains: If data exists in the repeat range,
  these values are written out too. The effect is that the rest of the rows is
  shifted down. This might destroy cell references. This doesn't happen when an
  existing ODS is read, but care has to be taken when generating new stuff.

- write_ods_buf and read_ods_buf implemented.

- add some missing namespaces.
- fix bug in Condition with string values.
- add range() to sheet.

# 0.7.0

- Add content validations.
    - Condition and ValueCondition allow for composition.
    - Validation supports all features except macros.
    - StyleMap uses ValueCondition now.
    - Validation can be set on SCell.
- CellRange, CellRef get absolute_xx() functions.

# 0.6.3

- Renamed the split functions to better match their functionality.
- Clippy brought up some issues.
- Moved Detached to the ds module. It's not a core concept.

# 0.6.2

- Rewriting an existing ODS was broken after the change to directly writing the
  ZIP.

# 0.6.1

- F'd up some tests with 0.6.0, and the new configuration features broke a lot.
  Fixed now.
- Reworked Config
    - Now the order of the configuration entries is kept between load/stores.
      Makes comparisons easier.
    - Added a few more configuration flags to be visible.
    - Added functions to sheet to set table-splits.

# 0.6.0

Breaking:

- Storing the WorkBook now needs a mutable reference.
- set_col_width doesn't need the WorkBook any longer.
- set_row_height doesn't need the WorkBook any longer.

Changes:

- Add basic row repeat functionality to Sheet. The subsequent row index will not
  be altered for now, and references etc must be updated manually. After writing
  and reading again the row indexes *will* be changed. So for now it's mostly
  useful to use this for the last row in the Sheet.
- The ODS content will no longer be written to a temp directory first and zipped
  later. Was a workaround a weird bug I couldn't locate. Testing now shows no
  problems.
- Styles without name get a name assigned when adding to the WorkBook.
- Length got an extra value Default.
- Sheet.set_col_width and Sheet.set_row_height now can work without the WorkBook
  Parameter. The necessary style modifications are applied when storing the
  workbook.
  *Storing the WorkBook now needs a mutable reference to make this possible.*
- Sheets can now be detached from the workbook and leave a placeholder behind to
  be reattached later. This makes it easier to modify the WorkBook and a Sheet
  at the same time.
- AttrMap2Trait removed, not helpful after the style reorg.
- settings.xml is now parsed. A subset of the Settings can be accessed via
  WorkBook.config() and Sheet.config().

# 0.5.2

- Add useability for XmlTag and TextTag.
- Implement a few standard TextTag Wrappers for common text elements.

# 0.5.1

- Split up the unwieldy PageLayout into PageStyle and MasterPage, and add an
  example how to use those bastards.

# 0.5.0
- Major reorg of styles. Replaced Styles with separate CellStyle, ColStyle,
  RowStyle etc.
- Create CellStyleRef, ColStyleRef, etc to be used when relating to styles.
  Should add a little bit of safety here.
- Introduced submodules for style: stylemap, tabstop, units

# 0.4.2

- Allow the ODS version to be specified. This adds support for ODS 1.3. -- Default version set to 1.3.

# 0.4.1 

- Refine usage of Style::cell(), cell_mut(), table(), table_mut(), col(), col_mut(),
  row(), row_mut(). Assert the correct style-family for access to these Attributes.

- Bug when writing empty lines, used wrong row-style.

- No row/column styles are written if they are beyond the range of the maximum
  used cell. This is a desired behaviour. To make it easier there is a
  Value-Conversion from '()' to an empty cell-value.

