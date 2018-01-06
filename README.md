[![CircleCI](https://circleci.com/gh/brendanlong/ocaml-xlsx-parser.svg?style=shield)](https://circleci.com/gh/brendanlong/ocaml-xlsx-parser)

`xlsx_parser` is an OCaml library for reading data from XLSX (Excel)
files. The goal is to eventually be pure-OCaml and relatively
performant, with Async and Lwt support, but currently camlzip makes
the interface pretty awkward (can't do things in memory, the XLSX
file has to exist on the filesystem; can't be async).

**Warning: Check the build status above. This is probably not ready
to use yet.**
