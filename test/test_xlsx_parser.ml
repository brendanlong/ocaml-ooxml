open Base
open OUnit2
open Printf

(* Test documents come from OpenOffice's test suite:
   https://www.openoffice.org/sc/testdocs/

   I downloaded the "XML" versions of each document, then converted them to
   CSV using OpenOffice. The tests then compare our input of the XLSX
   document to the CSV.

   Note: OpenOffice seems to round floats when exporting CSV, so I had to
   manually edit the CSV's to have the literal float values in the XLSX file. *)

let printer xlsx =
  Xlsx_parser.sexp_of_t xlsx
  |> Sexp.to_string_hum

let test_openoffice _ =
  let expect =
    let name = "Sheet1" in
    let rows =
      Csv.load ~strip:false "test/files/cells_import_xml_12.csv"
      |> Csv.to_array
      |> Array.map ~f:Array.to_list
      |> Array.to_list
    in
    [ { Xlsx_parser.name ; rows } ]
  in
  Xlsx_parser.read_file "test/files/cells_import_xml_12.xlsx"
  |> assert_equal ~printer expect

let () =
  [ "test_openoffice" >:: test_openoffice ]
  |> test_list
  |> run_test_tt_main
