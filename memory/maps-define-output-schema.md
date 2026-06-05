---
name: maps-define-output-schema
description: Collaborator stages expected-output files (e.g. tests/maps/*.xml) to define a feature's intended schema before asking for implementation
metadata:
  type: project
---

When implementing a new output format/feature here, check `tests/` for pre-staged
reference files that encode the *intended* shape before designing your own. For
the `--xml` feature, `tests/maps/data.xml` and `tests/maps/onBoarding.xml`
(staged hours before the request, with empty values) defined the schema:
root = camel-cased sheet name, header row → field element names
(`Last Name` → `<Last_Name>`, `FT/PT` → `<FT_PT>`), each data row → a `<row>`
record. Paired fixtures live in `tests/files/xml-*-book.xlsx`.

**Why:** The branch.prompt.md spec was terse ("output the mapped table of a
sheet as XML") and my first implementation (generic `<table>/<row>/<cell>` grid)
was wrong; the staged maps revealed the real field-mapped record schema.

**How to apply:** Before coding a format/output feature, grep the repo for
recently-added sample/expected/`maps/` files and match them exactly (filling in
real values where the map shows empty placeholders). The collaborator
communicates intent through staged artifacts, not just prose.
