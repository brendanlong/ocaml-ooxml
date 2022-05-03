open Base
open OUnit2
open Printf

(* Test documents come from OpenOffice's test suite:
   https://www.openoffice.org/sc/testdocs/

   I downloaded the "XML" versions of each document, then converted them to
   CSV using OpenOffice. The tests then compare our input of the XLSX
   document to the CSV. *)
type t =
  { name : string
  ; rows : string list list }

let pp_diff fmt (exp, real) =
  if String.(exp.name <> real.name) then
    Caml.Format.fprintf fmt "Expected sheet %s but got sheet %s" exp.name real.name
  else
    Caml.Format.fprintf fmt "Sheet: %s\n" exp.name;
  let print_opt_row diff_mark row =
    Option.iter row ~f:(fun row ->
      String.concat ~sep:"," row
      |> Caml.Format.fprintf fmt "%c %s\n" diff_mark)
  in
  let rec diff_lists = function
    | [], [] -> ()
    | exp, real ->
      let exp_row = List.hd exp in
      let real_row = List.hd real in
      if Option.equal (List.equal String.equal) exp_row real_row then
        print_opt_row ' ' exp_row
      else begin
        print_opt_row '-' exp_row;
        print_opt_row '+' real_row
      end;
      diff_lists
        (List.tl exp |> Option.value ~default:[],
         List.tl real |> Option.value ~default:[])
  in
  diff_lists (exp.rows, real.rows)

let sort_sheets =
  List.sort ~compare:(fun a b ->
    String.compare a.name b.name)

let make_test (file_name, sheet_names) =
  let normalize_sheets sheets =
    (* Fix sheet order and remove whitespace at the end of lines or sheets,
       since we don't care about either of these things *)
    sort_sheets sheets
    |> List.map ~f:(fun { name ; rows } ->
      let rows =
        List.rev_map rows ~f:(fun row ->
          List.rev row
          |> List.drop_while ~f:String.is_empty
          |> List.rev)
        |> List.drop_while ~f:List.is_empty
        |> List.rev
      in
      { name ; rows })
  in
  let sheet_names = List.sort ~compare:String.compare sheet_names in
  file_name >::
    fun _ ->
      let expect =
        List.map sheet_names ~f:(fun sheet_name ->
          let rows =
            sprintf "files/%s_%s.csv" file_name sheet_name
            |> Csv.load ~strip:false
            |> Csv.to_array
            |> Array.map ~f:Array.to_list
            |> Array.to_list
          in
          { name = sheet_name ; rows })
        |> normalize_sheets
      in
      Easy_xlsx.read_file (sprintf "files/%s.xlsx" file_name)
      |> List.map ~f:(fun { Easy_xlsx.name ; rows } ->
        let rows =
          List.map rows ~f:(List.map ~f:Easy_xlsx.Value.to_string) in
        { name ; rows })
      |> normalize_sheets
      |> List.iter2_exn expect ~f:(fun expect actual ->
        assert_equal ~pp_diff expect actual)

let () =
  [ "autofilter_import_xml_12", [ "Discrete"; "Top10"; "Custom"; "Advanced1"; "Advanced2" ]
  ; "chart_3dsettings_import_xml_12", [ "Rotation"; "Elevation"; "Perspective"; "Settings"; "SourceData" ]
  ; "chart_axis_import_xml_12", [ "Axis Line"; "Axis Labels"; "Tick Marks"; "Gridlines"; "SourceData" ]
  ; "hyperlink_import_xml_12", [ "Sheet1"; "Sheet2"; "Sheet'!" ]
  ; "oleobject_import_xml_12", [ "OLE"; "SourceData" ]
  ; "scenarios_import_xml_12", [ "Ranges"; "Settings"; "Name"; "RefCheck" ]
  ; "sheetprotection_import_xml_12", [ "Sheet1"; "Sheet2"; "Sheet3"; "Sheet4" ] ]
  |> List.map ~f:make_test
  |> test_list
  |> run_test_tt_main
