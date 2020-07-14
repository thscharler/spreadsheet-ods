![crates.io](https://img.shields.io/crates/v/spreadsheet-ods.svg)(https://crates.io/crates/spreadsheet-ods)
![Documentation](https://docs.rs/spreadsheet_ods/badge.svg)(https://docs.rs/spreadsheet_ods) 


spreadsheet-ods - Read and write ODS files
====

This crate can read and write back ODS spreadsheet files. 

Not all of the specification is implemented yet. And there are parts for 
which there is no public API, but which are preserved as raw xml. More 
details in the documentation.

## Usage

Add the following to your `Cargo.toml`:

```toml
[dependencies]
spreadsheet-ods = "0.4"
```

## Features

* `use_decimal`: Add conversions for rust_decimal. Internally the values are
  stored as f64 nonetheless.

* `dump_xml`: For debugging only.
* `dump_unused`: For debugging only. Writes out all xml tags that are not 
   processed.
* `indent_xml`: For debugging only. Pretty prints the generated xml. 
* `check_xml`: For debugging only: Checks for valid xml structure when writing.


## License

This project is licensed under either of

* [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0)
  ([LICENSE-APACHE](LICENSE-APACHE))

* [MIT License](https://opensource.org/licenses/MIT)
  ([LICENSE-MIT](LICENSE-MIT))

at your option.

## Contributing

I welcome all people who want to contribute.  
