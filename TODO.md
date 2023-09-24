Draft for 1.0

For what it's worth, I declare this version 1.0.

It covers the core functionality to read and write ODS files, and to expose
the functionality in it's api.

I'm quite satisfied with the current api, so now is a good time to stabilize
what is.

FUTURE

- On the spreadsheet level still missing are scripts, tracked-changes,
  variable-decls, sequence-decls, user-field-decls, dde-connection-decls,
  calculation-settings, label-ranges, named-expressions, database-ranges,
  data-pilot-tables, consolidation and dde-links.
  Anyway they are conserved during a read/write cycle.
- On the single table level still missing are dde-source, scenario, forms,
  shapes, named-expressions.
  They are also conserved during a read/write cycle.
- There is also no current plan to add charts and drawings.
- Support to parse/create formulas will be bigger projects, and I will rather
  support them with additional crates. There are some blueprints lying around
  but they need more work.
