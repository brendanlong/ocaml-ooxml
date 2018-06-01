[![CircleCI](https://circleci.com/gh/brendanlong/ocaml-ooxml.svg?style=shield)](https://circleci.com/gh/brendanlong/ocaml-ooxml)
[![Coverage Status](https://coveralls.io/repos/github/brendanlong/ocaml-ooxml/badge.svg?branch=master)](https://coveralls.io/github/brendanlong/ocaml-ooxml?branch=master)

The repo contains three libraries for reading data from Microsoft's document
formats ("Office Open XML").

  - `open_packaging` parses Office Open XML's "Open Packaging Conventions"
    (the container format for all Microsoft Office documents)
  - `spreadsheetml` parses the XML data in SpreadsheetML (i.e. Excel's XLSX
    format)
  - `easy_xlsx` reads XLSX documents, applies the formatting in the document,
    and returns the result as a list of sheets (with a sheet name and then
    a `string list list` of data). The goal of this library is to give the
    same results as if you exported the XLSX file to CSV and then read it with
    a standard CSV reader.

`open_packaging` and `spreadsheetml` are relatively safe to use but incomplete
(it should be obvious what they're mising -- if a field doesn't exist, I
haven't got to it yet). Everything that does exist should be parsed properly.

`easy_xlsx` is in very early stages. It should properly give read XLSX files
and output correct types, but the SpreadsheetML spec doesn't list all of the
built-in format strings, so some types may not be handled correctly. At the
moment, `easy_xlsx` will bail out in any case where it can't understand the
formatting, although I'd be open to patches to make this optional.

[API documentation](https://brendanlong.github.io/ocaml-ooxml/doc/)

## Building

Install dependencies:

```
opam pin add -n easy_xlsx .
opam depext easy_xlsx
opam install --deps-only easy_xlsx
```

Then build:

```
make
```

You can run the tests if you want:

```
make test
```

The `Makefile` is just a thin wrapper around jbuilder, so you can use
jbuilder commands too if you prefer.

## Helping

If you want to help with this, create an issue with what you'd like to work
on, and mention it if you need me to help with anything (point you at the
relevant spec, give my opinion on approaches, etc.).

Some things I could use help with:

  - Implement more of the specs. See [ECMA 376](https://www.ecma-international.org/publications/standards/Ecma-376.htm).
    The editions only contain changes, so most of what's interesting is in
    the 1st edition. Part 2 is the most interesting for the Open Packaging
    Conventions and Part 4 has specifics for SpreadsheetML (or the other office
    formats if you want to start a library for them).
  - Add more tests. Right now there's a set of extremely high level tests
    that we get the sound output as OpenOffice's CSV export, but being so high
    level means a lot of things are completely untested until we're 100%
    finished, which isn't a good situation to be in. It's easy to have a typo
    when implementing this spec, so I'd like to aim for 100% test coverage.
  - The cell formatting needs a lot of work. It's what I'm likely to work on
    next, but if you think you can do it first I'd be happy to let someone
    else do it (let me know if you start working on this though, so we don't
    duplicate effort).
  - [Make CamlZip support a string or bigstring interface](https://github.com/xavierleroy/camlzip/pull/7).
    Right now we can only read actual files on the filesystem since that's the
    interface CamlZip gives us. I'd like to be able to read data from memory.
  - Make a pure-OCaml ZIP library. Then we won't need any system dependencies
    and should work with `js_of_ocaml` too.
