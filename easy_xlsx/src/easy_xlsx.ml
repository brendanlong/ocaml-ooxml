open Core_kernel
open Stdint
open Spreadsheetml

let zip_entry_to_xml zip name =
  Zip.find_entry zip name
  |> Zip.read_entry zip
  |> Xml.parse_string

type sheet =
  { name : string
  ; rows : string list list }
    [@@deriving compare, sexp]

type t = sheet list [@@deriving compare, sexp]

let format_cell ~styles ~shared_strings
    ({ Worksheet.Cell.style_index ; data_type ; _ } as cell) =
  let str = Worksheet.Cell.to_string ~shared_strings cell in
  match data_type with
  | Number ->
    if str = "" then
      ""
    else
      let add_commas s =
        String.split s ~on:'.'
        |> (function
        | i :: p ->
          (String.to_list_rev i
           |> List.groupi ~break:(fun i _ _ -> i mod 3 = 0)
           |> List.map ~f:List.rev
           |> List.map ~f:String.of_char_list
           |> List.rev
           |> String.concat ~sep:",")
          :: p
        | l -> l)
        |> String.concat ~sep:"."
      in
      let n = float_of_string str in
      styles.(Uint32.to_int style_index)
      |> Spreadsheetml.Styles.Format.number_format_id
      |> Option.map ~f:Uint32.to_int
      |> Option.value ~default:0
      |> (function
      | 1 ->
        Float.iround_exn n
        |> Int.to_string
      | 2 ->
        sprintf "%.2f" n
      | 3 ->
        Float.iround_exn n
        |> Int.to_string
        |> add_commas
      | 4 ->
        sprintf "%.2f" n
        |> add_commas
      | 9 ->
        n *. 100.
        |> Float.iround_exn
        |> sprintf "%d%%"
      | 10 ->
        n *. 100.
        |> sprintf "%.2f%%"
      | 11 ->
        sprintf "%.2E" n
      (* TODO: 12 is defined as # ?/?, ex: 1234 4/7 *)
      (* TODO: 13 is defined as # ??/??, ex: 1234 46/81 *)
      | 14 ->
        Date.create_exn ~y:1899 ~m:Month.Dec ~d:30
        |> (fun d -> Date.add_days d (Float.iround_exn ~dir:`Down n))
        |> (fun d ->
          let month = Date.month d |> Month.to_int in
          let day = Date.day d in
          let year = Date.year d in
          sprintf "%d/%d/%d" month day year)
      | 0 | _ ->
        if Float.(round_down n = n) then
          Float.to_int n
          |> sprintf "%d"
        else
          Float.to_string n)
  | _ -> str

let read_file filename =
  let zip = Zip.open_in filename in
  Exn.protect ~f:(fun () ->
    let shared_strings =
      zip_entry_to_xml zip "xl/sharedStrings.xml"
      |> Shared_string_table.of_xml
      |> List.to_array
    in
    let sheets =
      zip_entry_to_xml zip "xl/workbook.xml"
      |> Workbook.of_xml
      |> Workbook.sheets
    in
    let styles =
      zip_entry_to_xml zip "xl/styles.xml"
      |> Styles.of_xml
      |> Styles.cell_formats
      |> Array.of_list
    in
    let rel_map =
      zip_entry_to_xml zip "xl/_rels/workbook.xml.rels"
      |> Open_packaging.Relationships.of_xml
      |> List.map ~f:(fun { Open_packaging.Relationship.id ; target ; _ } ->
        id, target)
      |> String.Map.of_alist_exn
    in
    List.map sheets ~f:(fun { Workbook.Sheet.name ; id ; _ } ->
      let rows =
        let target = Map.find_exn rel_map id in
        let path = sprintf "xl/%s" target in
        let { Worksheet.columns ; rows } =
          Zip.find_entry zip path
          |> Zip.read_entry zip
          |> Xml.parse_string
          |> Worksheet.of_xml
        in
        let num_cols =
          columns
          |> List.map ~f:Worksheet.Column.max
          |> List.map ~f:Uint32.to_int
          |> List.max_elt ~cmp:Int.compare
          |> Option.value ~default:0
        in
        let row_map =
          rows
          |> List.map ~f:(fun { Worksheet.Row.row_index ; cells ; _ } ->
            let index =
              Option.value_exn ~here:[%here] row_index
              |> Uint32.to_int
            in
            let cell_map =
              List.map cells ~f:(fun cell ->
                Worksheet.Cell.column cell, cell)
              |> Int.Map.of_alist_exn
            in
            let cells =
              Map.max_elt cell_map
              |> Option.map ~f:(fun (max, _) ->
                List.init (max + 1) ~f:(fun i ->
                  Map.find cell_map i
                  |> Option.value ~default:Worksheet.Cell.default))
              |> Option.value ~default:[]
            in
            index - 1, cells)
          |> Int.Map.of_alist_exn
        in
        let n =
          Map.keys row_map
          |> List.max_elt ~cmp:Int.compare
          |> Option.map ~f:((+) 1)
          |> Option.value ~default:0
        in
        List.init n ~f:Fn.id
        |> List.map ~f:(fun i ->
          let row =
            Map.find row_map i
            |> Option.value ~default:[]
          in
          let missing_cols = num_cols - List.length row in
          if missing_cols > 0 then
            row @ List.init ~f:(Fn.const Worksheet.Cell.default) missing_cols
          else
            row)
        |> List.map ~f:(List.map ~f:(format_cell ~styles ~shared_strings))
      in
      { name ; rows }))
    ~finally:(fun () -> Zip.close_in zip)
