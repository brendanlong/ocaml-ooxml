(lang dune 3.0)

(name easy_xlsx)

(generate_opam_files true)

(source
 (github brendanlong/ocaml-ooxml))

(authors "Brendan Long <self@brendanlong.com>")

(maintainers "Brendan Long <self@brendanlong.com>")

(license Unlicense)

(documentation https://brendanlong.github.io/ocaml-ooxml/doc)

(package
 (name open_packaging)
 (synopsis
  "A library for parsing Microsoft's Open Packaging Specification")
 (description
  "A library for parsing Microsoft's Open Packaging Specification (most commonly used for Microsoft Office's \"Office Open XML\")")
 (depends
  (ocaml (>= 5.2.1))
  base
  ppx_jane
  sexplib
  stdint
  xml-light))

(package
 (name spreadsheetml)
 (synopsis "A library for parsing SpreadsheetML (used in Microsoft Excel files)")
 (description
  "A library for parsing SpreadsheetML (used in Microsoft Excel files)")
 (depends
  (ocaml (>= 5.2.1))
  open_packaging
  ppx_jane))

(package
 (name easy_xlsx)
 (synopsis "A library to easily read XLSX files into a simpler format")
 (description "A library to easily read XLSX files into a simpler format")
 (depends
  (ocaml (>= 5.2.1))
  camlzip
  ppx_jane
  ptime
  spreadsheetml
  (csv :with-test)
  (ounit2 :with-test)))
