open Base
open Printf

let find_attr (attrs : (string * string) list) (name_to_find : string) =
  List.find_map_exn attrs ~f:(fun (name, value) ->
    if String.(name = name_to_find) then
      Some value
    else
      None)

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
      let x_window = find_attr attrs "xWindow" |> Int.of_string in
      let y_window = find_attr attrs "yWindow" |> Int.of_string in
      let window_width = find_attr attrs "windowWidth" |> Int.of_string in
      let window_height = find_attr attrs "windowHeight" |> Int.of_string in
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
      let name = find_attr attrs "name" in
      let id = find_attr attrs "sheetId" |> Int.of_string in
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
      |> let open Xml in function
      | Element ("sst", _, children) ->
        List.map children ~f:(function
          | Element ("si", _, [ Element ("t", _, []) ]) -> ""
          | Element ("si", _, [ Element ("t", _, [ PCData str ]) ]) -> str
          | el ->
            failwithf "Unexpected shared string element %s"
              (Xml.to_string el) ())
        |> List.to_array
      | _ -> assert false
    in
    let sheets =
      Workbook_meta.of_zip zip
      |> Workbook_meta.sheets
    in
    List.map sheets ~f:(fun { Sheet_meta.name ; id } ->
      let rows =
        let path = sprintf "xl/worksheets/sheet%d.xml" id in
        Zip.find_entry zip path
        |> Zip.read_entry zip
        |> Xml.parse_string
        |> let open Xml in function
        | Element ("worksheet", _, children) ->
          List.find_exn children ~f:(function
            | Element ("sheetData", _, _) -> true
            | _ -> false)
          |> Xml.children
          |> List.filter_map ~f:(function
          | Element ("row", _, cells) ->
            List.filter_map cells ~f:(function
              | Element ("c", attrs, [ Element ("v", _, [ PCData v ]) ]) ->
                let is_shared = List.exists attrs ~f:(function
                  | "t", "s" -> true
                  | _ -> false)
                in
                if is_shared then
                  let i = Int.of_string v in
                  Some shared_strings.(i)
                else
                  Some v
              | _ -> None)
            |> Option.return
          | _ -> None)
        | _ -> assert false
      in
      { name ; rows }))
    ~finally:(fun () -> Zip.close_in zip)
