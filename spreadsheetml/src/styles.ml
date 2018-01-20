(* 3.8 Styles *)
open Core_kernel
open Stdint
open Utils

module Format = struct
  (* 3.8.45 xf (Format) *)
  (* TODO: Add <alignment> and <protection> children *)
  type t =
    { apply_alignment : bool option
    ; apply_border : bool option
    ; apply_fill : bool option
    ; apply_font : bool option
    ; apply_number_format : bool option
    ; apply_protection : bool option
    ; border_id : uint32 option
    ; fill_id : uint32 option
    ; font_id : uint32 option
    ; number_format_id : uint32 option
    ; pivot_button : bool
    ; quote_prefix : bool
    ; format_id : uint32 option }
      [@@deriving fields, sexp_of]

  let default =
    { apply_alignment = None
    ; apply_border = None
    ; apply_fill = None
    ; apply_font = None
    ; apply_number_format = None
    ; apply_protection = None
    ; border_id = None
    ; fill_id = None
    ; font_id = None
    ; number_format_id = None
    ; pivot_button = false
    ; quote_prefix = false
    ; format_id = None }

  let of_xml = function
    | Xml.Element ("xf", attrs, _) ->
      List.fold attrs ~init:default ~f:(fun acc -> function
        | "applyAlignment", v ->
          { acc with apply_alignment = Some (bool_of_xsd_boolean v) }
        | "applyBorder", v ->
          { acc with apply_border = Some (bool_of_xsd_boolean v) }
        | "applyFill", v ->
          { acc with apply_fill = Some (bool_of_xsd_boolean v) }
        | "applyFont", v ->
          { acc with apply_font = Some (bool_of_xsd_boolean v) }
        | "applyNumberFormat", v ->
          { acc with apply_number_format = Some (bool_of_xsd_boolean v) }
        | "applyProtection", v ->
          { acc with apply_protection = Some (bool_of_xsd_boolean v) }
        | "borderId", v ->
          { acc with border_id = Some (Uint32.of_string v) }
        | "fillId", v ->
          { acc with fill_id = Some (Uint32.of_string v) }
        | "fontId", v ->
          { acc with font_id = Some (Uint32.of_string v) }
        | "numFmtId", v ->
          { acc with number_format_id = Some (Uint32.of_string v) }
        | "pivotButton", v ->
          { acc with pivot_button = bool_of_xsd_boolean v }
        | "quotePrefix", v ->
          { acc with quote_prefix = bool_of_xsd_boolean v }
        | "xfId", v ->
          { acc with format_id = Some (Uint32.of_string v) }
        | _ -> acc)
      |> Option.some
    | _ -> None
end

module Number_format = struct
  (* 3.8.30 numFmt (Number Format) *)
  type t =
    { id : uint32
    ; format : string }
      [@@deriving fields, sexp_of]

  let of_xml = function
    | Xml.Element ("numFmt", attrs, _) ->
      let id = ref None in
      let format = ref None in
      List.iter attrs ~f:(function
      | "numFmtId", v -> id := Some (Uint32.of_string v)
      | "formatCode", v -> format := Some v
      | _ -> ());
      let id = require_attribute "numFmt" "numFmtId" !id in
      let format = require_attribute "numFmt" "formatCode" !format in
      Some { id ; format }
    | _ -> None
end

type t =
  { cell_formats : Format.t list
  ; formatting_records : Format.t list
  ; number_formats : Number_format.t list }
    [@@deriving fields, sexp_of]

let empty =
  { cell_formats = []
  ; formatting_records = []
  ; number_formats = [] }

let of_xml =
  expect_element "styleSheet" (fun _ ->
    List.fold ~init:empty ~f:(fun acc ->
      let open Xml in
      function
      | Element ("cellStyleXfs", _, children) ->
        let formatting_records = List.filter_map children ~f:Format.of_xml in
        { acc with formatting_records }
      | Element ("cellXfs", _, children) ->
        let cell_formats = List.filter_map children ~f:Format.of_xml in
        { acc with cell_formats }
      | Element ("numFmts", _, children) ->
        let number_formats =
          List.filter_map children ~f:Number_format.of_xml in
        { acc with number_formats }
      | _ -> acc))
