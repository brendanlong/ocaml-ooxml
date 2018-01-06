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

module Workbook_view = struct
  type t =
    { x_window : int
    ; y_window : int
    ; window_width : int
    ; window_height : int }
      [@@deriving compare, fields, sexp]

  let of_xml =
    let open Xml in
    function
    | Element ("workbookView", attrs, _) ->
      let x_window = find_attr_exn attrs "xWindow" |> Int.of_string in
      let y_window = find_attr_exn attrs "yWindow" |> Int.of_string in
      let window_width = find_attr_exn attrs "windowWidth" |> Int.of_string in
      let window_height = find_attr_exn attrs "windowHeight" |> Int.of_string in
      Some { x_window ; y_window ; window_width ; window_height }
    | _ -> None
end

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
    { book_views : Workbook_view.t list
    ; sheets : Sheet_meta.t list }
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
      let book_views =
        List.find_map children ~f:(function
        | Element ("bookViews", _, views) ->
          Some (List.filter_map views ~f:Workbook_view.of_xml)
        | _ -> None)
        |> Option.value ~default:[]
      in
      { book_views ; sheets }
    | _ -> assert false
end

type row = string list [@@deriving compare, sexp]

type sheet =
  { name : string
  ; rows : row list }
    [@@deriving compare, sexp]

type t = sheet list [@@deriving compare, sexp]

let read_file filename =
  let zip = Zip.open_in filename in
  Exn.protect ~f:(fun () ->
    let shared_strings =
      Zip.find_entry zip "xl/sharedStrings.xml"
      |> Zip.read_entry zip
      |> Xml.parse_string
      |> find_elements ~path:[ "sst" ; "si" ]
      |> List.map ~f:(function
      | Element ("t", _, [ PCData str] ) -> str
      | el ->
        failwithf "Unexpected shared string element %s"
          (Xml.to_string el) ())
      |> List.to_array
    in
    let sheets =
      Workbook_meta.of_zip zip
      |> Workbook_meta.sheets
    in
    List.map sheets ~f:(fun { Sheet_meta.name ; id } ->
      let rows =
        let path = sprintf "xl/worksheets/sheet%d.xml" id in
        let root =
          Zip.find_entry zip path
          |> Zip.read_entry zip
          |> Xml.parse_string
        in
        let num_cols =
          find_elements root ~path:[ "worksheet" ; "cols" ]
          |> List.filter_map ~f:(function
          | Element ("col", attrs, _) ->
            find_attr attrs "max" |> Option.map ~f:Int.of_string
          | _ -> None)
          |> List.max_elt ~cmp:Int.compare
          |> Option.value ~default:0
        in
        let row_map =
          let open Xml in
          find_elements root ~path:[ "worksheet" ; "sheetData" ]
          |> List.filter_map ~f:(function
          | Element ("row", attrs, cells) ->
            let index = find_attr_exn attrs "r" |> Int.of_string in
            let row =
              List.map cells ~f:(fun cell ->
                match cell with
                | Element ("c", attrs, children) ->
                  let t = List.find_map attrs ~f:(function
                    | "t", value -> Some value
                    | _ -> None)
                  in
                  (match t with
                  | Some "inlineStr" ->
                    find_elements cell ~path:[ "c" ; "is" ; "t" ]
                    |> List.find_map ~f:(function
                    | PCData v -> Some v
                    | _ -> None)
                    |> Option.value ~default:""
                  | None | Some "s" | _->
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
                | _ -> "")
            in
            (* Rows at 1-indexed, convert to 0-indexed *)
            Some (index - 1, row)
          | _ -> None)
          |> Map.Using_comparator.of_alist_exn ~comparator:Int.comparator
        in
        let n =
          Map.keys row_map
          |> List.max_elt ~cmp:Int.compare
          |> Option.value ~default:0
        in
        let default_row = List.init num_cols ~f:(Fn.const "") in
        List.init (n + 1) ~f:Fn.id
        |> List.map ~f:(fun i ->
          Map.find row_map i
          |> Option.value ~default:default_row)
      in
      { name ; rows }))
    ~finally:(fun () -> Zip.close_in zip)
