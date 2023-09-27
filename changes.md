# 0.17.0

NEW

- Allow access to meta.xml data.
- Allow access to manifest.xml.
- Add support for row and column groupings.

CHANGED

- Rewrote the XMLWriter to cause less allocations. Mixed results, but nicer API.
- Removed support for header-rows/header-columns. This is only used for Writer
  not for Spreadsheet.
- Datetime values can have a trailing "Z".
- Basic support for ruby-styles.
- Add missing iterators for WorkBook content.

# 0.16.1

- Add PartialEq for Value and dependencies.
- Add WorkBook::iter_sheets(), iter_row_styles(), iter_col_styles(),
  iter_cell_styles().
- Bump dependencies.
- Reexport color-rs crate as spreadsheet_ods::color. It seems this is often with
  the defunct "color" crate.
- Fixed a compile-error PR#44

# 0.16.0

- New ValueStyleMap for use in ValueFormat*.
- base_cell is optional even for CellStyle stylemaps.
- ValueCondition has to use 'value()'

- read_ods_from() and write_ods_to() for Read/Write traits.

# 0.15.0

- It was an error to assume that currency values use an ISO code for
  the currency string. Removed the optimization and use a String again.

- number-rows-repeated a million times. Can be found for the last or the
  second to last row. If the row is overwritten with actual data and
  opened in LibreOffice this results to a real memory stress test.
  Any repeat count of more than 1000 for the last two rows are now ignored.

- Sheet::split_col_header() and split_row_header() now split after
  the given row/column.
- Add as_*64, as_*16, as_*8 conversions for Value.

- Bug: Default number-format should set min-integer-digits to 1. Fixed.
- Bug: LibreOffice uses dates like 0000-00-00. Fixed.
- Bug: embedded-text in format broke the parser. Removed that part for now
  and ignore this tag.
- Bug: Parsing sheet-names failed with the new reference parser. Fixed.

- Update dependencies

# 0.14.0

- Undo spreadsheet-ods-cellref. Was a reasonable start, but didn't work out
  as expected.
- Instead use a splinter of a parser for OpenFormula I'm working on separately
  for cellref parsing.
- This means
    - Cell-references now can contain external references via an IRI.
    - Cell-ranges can span more than one table.
    - Colranges and Rowranges have IRI, from-table and to-table now too.

# 0.13.0

- Upgrade mktemp to latest.
- Extracted cell references to a separate crate spreadsheet-ods-cellref.
    - The parser has been rewritten with nom.
    - The fmt* functions are new too.
- CellRef
    - Add an IRI for references to external files.
- CellRange
    - Add an IRI for references to external files.
    - Add a to_table to allow ranges that span multiple sheets.
- ColRange, RowRange
    - Add an IRI for references to external files.
    - Add from_table and to_table.
    - Add from_col_abs, to_col_abs for fixed columns in ColRange.
    - Add from_row_abs, to_row_abs for fixed rows in RowRange.

# 0.12.1

- Upgrade icu_locid and quick_xml to latest.

# 0.12.0

BREAKING:

- ValueFormat is gone. Many, many functions had an annotation
  "can only be used when ...", which is not a good sign.
  So I split it up in one struct per ValueType (ValueFormatBoolean,
  ValueFormatNumber, ...). This allows for a clearer communication
  what is possible with each of them.

  Changing should be straightforward:

  Before:
  ```rust
    let mut v1 = ValueFormat::new_named("f1", ValueType::Number);
    v1.part_scientific().decimal_places(4).build();
    let v1 = wb.add_format(v1);
  ```

  After:
  ```rust
    let mut v1 = ValueFormatNumber::new_named("f1");
    v1.part_scientific().decimal_places(4).build();
    let v1 = wb.add_number_format(v1);
  ```

  The good news: I think I am happy now how ValueFormatXXX and XXXStyle work.
  I will keep them stable from now on.

CHANGES:

- create_loc_number_format_fixed, create_loc_time_interval_format where missing.
- HeaderFooter can contain multiple paragraphs of text. Works now.
- TextTag/XmlTag: Add functionality to work with Vec<XmlTag>.

# 0.11.1

- Minor fixes.

# 0.11.0

BREAKING:

Localization has been added via icu_locid. This leads to a few but central
breaks in the api.

- WorkBook::new() now needs a Locale. This obsoletes the call to
  create_default_styles()
  which never was really satisfying. The old behaviour can be had with
  WorkBook::new_empty()

- ValueFormat: set_country(), set_language(), set_script() were replaced with
  set_locale().
- ValueFormat: all the format_xxx() functions were a train-wreck and have been
  removed. They were only ever used to write the cell-content in a nicer way. A
  value
  that is immediately thrown away when the spreadsheet is openend. So I now
  write the
  same format that is used for the xxx-value attribute anyway.
- FormatPart: all new_xxx functions removed.

CHANGES:

- Overhauled ValueFormat.
    - All the ValueFormat::push_xxx were broken and missing attributes.
      As most of these attributes are optional these functions were replaced
      with new ValueFormat::part_xxx which return a builder for each pattern.

- Add icu_locid to the dependencies. Used where language/country/script
  attributes exist.
- Add locale module that contains localized default formats.
    - Available locales are behind feature-gates.
    - Needs ca 60 loc for a new locale.
    - Fallback available.
    - create_default_styles replaced with WorkBook::init_defaults and WorkBook::
      new_localized.

- Sheet::new() now always needs a name for the sheet.

- All the style attributes are crosschecked with the specification, and a lot
  of missing ones where added. I only excluded obviously obsolete ones and
  things that are out of scope.

- TableStyle::set_master_page_name() -> set_master_page()
- FontFaceDecl::new_with_name()

# 0.10.0

- Upgraded to edition 2021.
- Updated dependencies:
    - rust_decimal to 1.24
    - color_rs to 0.7
    - time to 0.3
    - zip to 0.6
    - removed criterion as dev-dependency.
- Parsing values implemented with nom and changed from str to &[u8] to safe on
  unnecessary utf8 conversion.
- Needed a lot of read buffers for each xml hierarchy level. Keep them around
  and reuse them.
- set_row_repeat must not be 0. Panics if so. This doesn't solve all
  problems with set_row_repeat, there is still some spurious repeat on the
  last row.
- Content validation was broken.

# 0.9.0

- Throw away SCell. This was used for internal storage and as part of the API.
  Split this into the internal CellData and the API CellContent for a copy
  of the cell data and CellContentRef for references to the data.
  This allows for a possible future rearrangement of the internal storage.

  cell_mut was removed, cell, add_cell, remove_cell, work with CellContent now.
  iter() and range() use CellContentRef.

- Throw away ucell. Uses u32 instead.

- Implement IntoIterator, iter(), range()

- Add CellSpan for ease of use.

- Changed layout of Value::Currency. The currency string is 3 bytes of ASCII,
  so a String is not necessary.

- read_table_cell and read_text_or_tag rewritten to use fewer copies of String
  data. Parsing cell-values works directly with the buffer data.

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

- Allow the ODS version to be specified. This adds support for ODS 1.3. --
  Default version set to 1.3.

# 0.4.1

- Refine usage of Style::cell(), cell_mut(), table(), table_mut(), col(),
  col_mut(),
  row(), row_mut(). Assert the correct style-family for access to these
  Attributes.

- Bug when writing empty lines, used wrong row-style.

- No row/column styles are written if they are beyond the range of the maximum
  used cell. This is a desired behaviour. To make it easier there is a
  Value-Conversion from '()' to an empty cell-value.

