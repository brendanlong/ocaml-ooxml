(* 3.3 Worksheets *)
open Core_kernel
open Stdint
open Utils

module Cell = struct
  module Type = struct
    (* 3.18.12 ST_CellType (Cell Type) *)
    type t =
      | Boolean
      | Error
      | Inline_string
      | Number
      | Shared_string
      | Formula_string
        [@@deriving sexp_of]

    let of_string = function
      | "b" -> Boolean
      | "e" -> Error
      | "inlineStr" -> Inline_string
      | "n" -> Number
      | "s" -> Shared_string
      | "str" -> Formula_string
      | str -> failwithf "Expected ST_CellType but got '%s'" str ()
  end

  module Value = struct
    (* 3.3.1.93 v (Cell Value) *)
    type t =
      | Inline of string
      | Rich_inline of Shared_string_table.String_item.t
      | Shared of uint32
          [@@deriving sexp_of]

    let to_string ~shared_strings = function
      | Inline v -> v
      | Rich_inline v ->
        Shared_string_table.String_item.to_string v
      | Shared n ->
        let n = Uint32.to_int n in
        shared_strings.(n)
        |> Shared_string_table.String_item.to_string

    let of_xml ~data_type el =
      let open Xml in
      match data_type, el with
      | _, Element ("is", _, children) ->
        let value =
          List.filter_map children ~f:Shared_string_table.Rich_text.of_xml in
        Some (Rich_inline value)
      | Type.Shared_string, Element ("v", _, children) ->
        Some (Shared (Uint32.of_string (expect_pcdata children)))
      | _, Element ("v", _, children) ->
        Some (Inline (expect_pcdata children))
      | _ -> None
  end

  (* 3.3.1.3 c (Cell) *)
  type t =
    { reference : string option
    ; formula : string option
    ; value : Value.t option
    ; cell_metadata_index : uint32
    ; show_phonetic : bool
    ; style_index : uint32
    ; data_type : Type.t
    ; value_metadata_index : uint32 }
      [@@deriving fields, sexp_of]

  let column { reference ; _ } =
    (Option.value_exn reference ~here:[%here]
     |> String.to_list
     |> List.take_while ~f:Char.is_alpha
     |> List.map ~f:Char.uppercase
     |> List.map ~f:(fun c -> Char.to_int c - Char.to_int 'A' + 1)
     |> List.fold ~init:0 ~f:(fun acc n -> acc * 26 + n))
    - 1

  let to_string ~shared_strings { value ; _ } =
    Option.map value ~f:(Value.to_string ~shared_strings)
    |> Option.value ~default:""

  let default =
    { reference = None
    ; formula = None
    ; value = None
    ; cell_metadata_index = Uint32.zero
    ; show_phonetic = false
    ; style_index = Uint32.zero
    ; data_type = Type.Number
    ; value_metadata_index = Uint32.zero }

  let of_xml =
    let open Xml in
    function
    | Element ("c", attrs, children) ->
      let with_attrs =
        List.fold attrs ~init:default ~f:(fun acc -> function
        | "r", v ->
          { acc with reference = Some v }
        | "s", v ->
          { acc with style_index = Uint32.of_string v }
        | "t", v ->
          { acc with data_type = Type.of_string v }
        | "cm", v ->
          { acc with cell_metadata_index = Uint32.of_string v }
        | "vm", v ->
          { acc with value_metadata_index = Uint32.of_string v }
        | "ph", v ->
          { acc with show_phonetic = bool_of_xsd_boolean v }
        | _ -> acc)
      in
      List.fold children ~init:with_attrs ~f:(fun acc -> function
      | Element ("f", _, children) ->
        { acc with formula = Some (expect_pcdata children) }
      | el ->
        let value = Value.of_xml ~data_type:with_attrs.data_type el in
        { acc with value = Option.first_some value acc.value })
      |> Option.some
    | _ -> None
end

module Row = struct
  (* 3.3.1.71 row (Row) *)
  type t =
    { cells : Cell.t list
    ; collapsed : bool
    ; custom_format : bool
    ; custom_height : bool
    ; hidden : bool
    ; height : float option
    ; outline_level : uint8
    ; show_phonetic : bool
    ; row_index : uint32 option
    ; style_index : uint32
    ; thick_bottom_border : bool
    ; thick_top_border : bool }
      [@@deriving fields, sexp_of]

  let default =
    { cells = []
    ; collapsed = false
    ; custom_format = false
    ; custom_height = false
    ; hidden = false
    ; height = None
    ; outline_level = Uint8.zero
    ; show_phonetic = false
    ; row_index = None
    ; style_index = Uint32.zero
    ; thick_bottom_border = false
    ; thick_top_border = false }

  let of_xml = function
    | Xml.Element ("row", attrs, children) ->
      let cells = List.filter_map children ~f:Cell.of_xml in
      List.fold attrs ~init:{ default with cells } ~f:(fun acc -> function
      | "collapsed", v ->
        { acc with collapsed = bool_of_xsd_boolean v }
      | "customFormat", v  ->
        { acc with custom_format = bool_of_xsd_boolean v }
      | "customHeight", v ->
        { acc with custom_height = bool_of_xsd_boolean v }
      | "hidden", v ->
        { acc with hidden = bool_of_xsd_boolean v }
      | "ht", v ->
        { acc with height = Some (float_of_string v) }
      | "outlineLevel", v ->
        { acc with outline_level = Uint8.of_string v }
      | "ph", v ->
        { acc with show_phonetic = bool_of_xsd_boolean v }
      | "r", v ->
        { acc with row_index = Some (Uint32.of_string v) }
      | "s", v ->
        { acc with style_index = Uint32.of_string v }
      | "thickBot", v ->
        { acc with thick_bottom_border = bool_of_xsd_boolean v }
      | "thickTop", v ->
        { acc with thick_top_border = bool_of_xsd_boolean v }
      | _ -> acc)
      |> Option.some
    | _ -> None
end

module Column = struct
  (* 3.3.1.12 col (Column Width & Formatting) *)
  type t =
    { best_fit : bool
    ; collapsed : bool
    ; custom_width : bool
    ; hidden : bool
    ; max : uint32
    ; min : uint32
    ; outline_level : uint8
    ; show_phonetic : bool
    ; default_style : uint32
    ; width : float option }
      [@@deriving fields, sexp_of]

  let of_xml = function
    | Xml.Element ("col", attrs, _) ->
      let best_fit = ref false in
      let collapsed = ref false in
      let custom_width = ref false in
      let hidden = ref false in
      let max = ref None in
      let min = ref None in
      let outline_level = ref Uint8.zero in
      let show_phonetic = ref false in
      let default_style = ref Uint32.zero in
      let width = ref None in
      List.iter attrs ~f:(function
      | "bestFit", v ->
        best_fit := bool_of_xsd_boolean v
      | "collapsed", v ->
        collapsed := bool_of_xsd_boolean v
      | "customWidth", v ->
        custom_width := bool_of_xsd_boolean v
      | "hidden", v ->
        hidden := bool_of_xsd_boolean v
      | "max", v ->
        max := Some (Uint32.of_string v)
      | "min", v ->
        min := Some (Uint32.of_string v)
      | "outlineLevel", v ->
        outline_level := Uint8.of_string v
      | "phonetic", v->
        show_phonetic := bool_of_xsd_boolean v
      | "style", v ->
        default_style := Uint32.of_string v
      | "width", v ->
        width := Some (float_of_string v)
      | _ -> ());
      let max = require_attribute "col" "max" !max in
      let min = require_attribute "col" "min" !min in
      Some { best_fit = !best_fit
           ; collapsed = !collapsed
           ; custom_width = !custom_width
           ; hidden = !hidden
           ; max
           ; min
           ; outline_level = !outline_level
           ; show_phonetic = !show_phonetic
           ; default_style = !default_style
           ; width = !width }
    | _ -> None
end

type t =
  { columns : Column.t list
  ; rows : Row.t list }
    [@@deriving fields, sexp_of]

let default =
  { columns = []
  ; rows = [] }

let of_xml =
  expect_element "worksheet" (fun _ ->
    let open Xml in
    List.fold ~init:default ~f:(fun acc -> function
    | Element ("cols", _, children) ->
      { acc with columns = List.filter_map children ~f:Column.of_xml }
    | Element ("sheetData", _, children) ->
      { acc with rows = List.filter_map children ~f:Row.of_xml }
    | _ -> acc))
