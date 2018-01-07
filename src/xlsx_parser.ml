open Base
open Printf

let find_attr (attrs : (string * string) list) (name_to_find : string) =
  List.find_map attrs ~f:(fun (name, value) ->
    if String.(name = name_to_find) then
      Some value
    else
      None)

let find_attr_exn attrs name_to_find =
  find_attr attrs name_to_find
  |> Option.value_exn ~here:[%here]

let col_of_cell_id id =
  (String.to_list id
   |> List.take_while ~f:Char.is_alpha
   |> List.map ~f:Char.uppercase
   |> List.map ~f:(fun c -> Char.to_int c - Char.to_int 'A' + 1)
   |> List.fold ~init:0 ~f:(fun acc n -> acc * 26 + n))
  - 1

let rec find_elements el ~path =
  match path with
  | [] -> [ el ]
  | search_tag :: tl ->
    match el with
    | Xml.Element (tag, _, children) when String.(tag = search_tag) ->
      List.concat_map children ~f:(find_elements ~path:tl)
    | _ -> []

let find_element_exn el ~path =
  find_elements el ~path
  |> List.hd_exn

module Sheet_meta = struct
  type t =
    { name : string
    ; id : int }
      [@@deriving compare, fields, sexp]

  let of_xml =
    let open Xml in
    function
    | Element ("sheet", attrs, _) ->
      let name = find_attr_exn attrs "name" in
      let id = find_attr_exn attrs "sheetId" |> Int.of_string in
      Some { name ; id }
    | _ -> None
end

module Workbook_meta = struct
  type t =
    { sheets : Sheet_meta.t list }
      [@@deriving compare, fields, sexp]

  let of_zip zip =
    let open Xml in
    Zip.find_entry zip "xl/workbook.xml"
    |> Zip.read_entry zip
    |> Xml.parse_string
    |> function
    | Element ("workbook", _, children) ->
      let sheets =
        List.find_map children ~f:(function
        | Element ("sheets", _, sheets) ->
          Some (List.filter_map sheets ~f:Sheet_meta.of_xml)
        | _ -> None)
        |> Option.value ~default:[]
      in
      { sheets }
    | _ -> assert false
end

module Column = struct
  type t =
    { min : int
    ; max : int }
      [@@deriving compare, fields, sexp]

  let of_xml = function
    | Xml.Element ("col", attrs, _) ->
      let min = find_attr_exn attrs "min" |> Int.of_string in
      let max = find_attr_exn attrs "max" |> Int.of_string in
      Some { min ; max }
    | _ -> None

  let of_worksheet_xml root =
    find_elements root ~path:[ "worksheet" ; "cols" ]
    |> List.filter_map ~f:of_xml
end

module Cell = struct
  type t = string [@@deriving compare, sexp]
end

module Row = struct
  type t =
    { index : int
    ; cells : Cell.t list }
      [@@deriving compare, fields, sexp]

  let of_xml ~shared_strings =
    let open Xml in
    function
    | Element ("row", attrs, cells) ->
      let index = find_attr_exn attrs "r" |> Int.of_string in
      let row =
        let cell_map =
          List.map cells ~f:(fun cell ->
            match cell with
            | Element ("c", attrs, children) ->
              let col =
                (* get the column number from the "r" attribute, which looks
                   like A1, B1, etc. *)
                find_attr_exn attrs "r"
                |> col_of_cell_id
              in
              let t = List.find_map attrs ~f:(function
                | "t", value -> Some value
                | _ -> None)
              in
              col, (match t with
              | Some "inlineStr" ->
                find_elements cell ~path:[ "c" ; "is" ; "t" ]
                |> List.find_map ~f:(function
                | PCData v -> Some v
                | _ -> None)
                |> Option.value ~default:""
              | _ ->
                find_elements cell ~path:[ "c" ; "v" ]
                |> List.find_map ~f:(function
                | PCData v ->
                  if Option.equal String.equal t (Some "s") then
                    let i = Int.of_string v in
                    Some shared_strings.(i)
                  else
                    Some v
                | _ -> None)
                |> Option.value ~default:"")
            | _ -> assert false)
          |> Map.Using_comparator.of_alist_exn ~comparator:Int.comparator
        in
        let n =
          Map.keys cell_map
          |> List.max_elt ~cmp:Int.compare
          |> Option.map ~f:((+) 1)
          |> Option.value ~default:0
        in
        List.init n ~f:Fn.id
        |> List.map ~f:(fun i ->
          Map.find cell_map i
          |> Option.value ~default:"")
      in
      (* Rows at 1-indexed, convert to 0-indexed *)
      Some { index = index - 1 ; cells = row }
    | _ -> None

  let of_worksheet_xml ~shared_strings root =
    find_elements root ~path:[ "worksheet" ; "sheetData" ]
    |> List.filter_map ~f:(of_xml ~shared_strings)
end

module Worksheet = struct
  type t =
    { columns : Column.t list
    ; rows : Row.t list }
      [@@deriving compare, fields, sexp]

  let of_xml ~shared_strings root =
    let columns = Column.of_worksheet_xml root in
    let rows = Row.of_worksheet_xml ~shared_strings root in
    { columns ; rows }
end

module Shared_strings = struct
  type t = string list [@@deriving compare, sexp]

  let of_xml root =
    find_elements ~path:[ "sst" ; "si" ] root
    |> List.map ~f:(function
    | (Element ("r", _, _) as el) ->
      find_elements el ~path:[ "r"; "t" ]
      |> List.find_map ~f:(function
      | PCData str -> Some str
      | _ -> None)
      |> Option.value ~default:""
    | Element ("t", _, []) -> ""
    | Element ("t", _, [ PCData str] ) -> str
    | el ->
      failwithf "Unexpected shared string element %s"
        (Xml.to_string el) ())
    |> List.to_array

  let of_zip zip =
    Zip.find_entry zip "xl/sharedStrings.xml"
    |> Zip.read_entry zip
    |> Xml.parse_string
    |> of_xml

end

type sheet =
  { name : string
  ; rows : string list list }
    [@@deriving compare, sexp]

type t = sheet list [@@deriving compare, sexp]

let read_file filename =
  let zip = Zip.open_in filename in
  Exn.protect ~f:(fun () ->
    let shared_strings = Shared_strings.of_zip zip in
    let sheets =
      Workbook_meta.of_zip zip
      |> Workbook_meta.sheets
    in
    List.map sheets ~f:(fun { Sheet_meta.name ; id } ->
      let rows =
        let path = sprintf "xl/worksheets/sheet%d.xml" id in
        let worksheet =
          Zip.find_entry zip path
          |> Zip.read_entry zip
          |> Xml.parse_string
          |> Worksheet.of_xml ~shared_strings
        in
        let num_cols =
          worksheet.Worksheet.columns
          |> List.map ~f:Column.max
          |> List.max_elt ~cmp:Int.compare
          |> Option.value ~default:0
        in
        let row_map =
          worksheet.Worksheet.rows
          |> List.map ~f:(fun { Row.index ; cells } ->
            index, cells)
          |> Map.Using_comparator.of_alist_exn ~comparator:Int.comparator
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
            row @ List.init ~f:(Fn.const "") missing_cols
          else
            row)
      in
      { name ; rows }))
    ~finally:(fun () -> Zip.close_in zip)
