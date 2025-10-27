[![crates.io](https://img.shields.io/crates/v/spreadsheet-ods.svg)](https://crates.io/crates/spreadsheet-ods)
[![Documentation](https://docs.rs/spreadsheet-ods/badge.svg)](https://docs.rs/spreadsheet_ods)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![License](https://img.shields.io/badge/license-APACHE-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)
[](https://tokei.rs/b1/github/thscharler/spreadsheet-ods)

# spreadsheet-ods - Read and write ODS files

This crate can read and write back ODS spreadsheet files.

Not all the specification is implemented yet.

But it covers everything if you

- just want to extract the data.
- need a round-trip. Any unparsed parts are kept as blobs or raw xml
  and are written again.

See the [todos]([changes.md](https://github.com/thscharler/spreadsheet-ods/blob/master/TODO.md)
) for the missing parts.

## Features

* `use_decimal`: Add conversions for rust_decimal. Internally the values are
  stored as f64 nonetheless.

* Locales
    * all_locales = [ "locale_de_AT", "locale_en_US", "locale_cs_CZ" ]
    * locale_de_AT
    * locale_en_US

  ... send me an issue if you need more ...

## License

This project is licensed under either of

* [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0)
  ([LICENSE-APACHE](LICENSE-APACHE))

* [MIT License](https://opensource.org/licenses/MIT)
  ([LICENSE-MIT](LICENSE-MIT))

at your option.

## Changes

[changes](https://github.com/thscharler/spreadsheet-ods/blob/master/changes.md)

## Contributing

I welcome all people who want to contribute.  
