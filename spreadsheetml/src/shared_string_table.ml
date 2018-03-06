open Base
open Base.Printf
open Stdint
open Utils

module Rich_text = struct
(* Note: We represent text and rich text the same internally. Unformatted
   text is just rich text with empty formatting options *)

  module Underline = struct
    (* http://www.datypic.com/sc/ooxml/t-ssml_ST_UnderlineValues.html *)
    type t =
      | Single
      | Double
      | Single_accounting
      | Double_accounting
      | None
        [@@deriving sexp_of]

    let of_string = function
      | "single" -> Single
      | "double" -> Double
      | "singleAccounting" -> Single_accounting
      | "doubleAccounting" -> Double_accounting
      | "none" -> None
      | str -> failwithf "Expected ST_Underline but got '%s'" str ()
  end

  module Vertical_align = struct
    (* http://www.datypic.com/sc/ooxml/t-ssml_ST_VerticalAlignRun.html *)
    type t =
      | Baseline
      | Superscript
      | Subscript
        [@@deriving sexp_of]

    let of_string = function
      | "baseline" -> Baseline
      | "superscript" -> Superscript
      | "subscript" -> Subscript
      | str -> failwithf "Expected ST_VerticalAlign but got '%s'" str ()
  end

  type t =
    { text : string
    ; bold : bool option
    ; charset : int32 option
    (* ; color : CT_Color *)
    ; condense : bool option
    ; extend : bool option
    ; font : string option
    ; font_family : int32 option
    ; font_size : float option (* CT_FontSize *)
    (* ; font_scheme : CT_FontScheme *)
    ; italic : bool option
    ; outline : bool option
    ; shadow : bool option
    ; strikethrough : bool option
    ; underline : Underline.t option
    ; vertical_align : Vertical_align.t option }
      [@@deriving fields, sexp_of]

  let empty =
    { text = ""
    ; bold = None
    ; charset = None
    ; condense = None
    ; extend = None
    ; font = None
    ; font_family = None
    ; font_size = None
    ; italic = None
    ; outline = None
    ; shadow = None
    ; strikethrough = None
    ; underline = None
    ; vertical_align = None }

  let of_xml =
    let open Xml in
    function
    | Element ("r", _, children) ->
      List.fold children ~init:empty ~f:(fun acc -> function
        | Element ("rPr", attrs, _) ->
          List.fold attrs ~init:acc ~f:(fun acc -> function
            | "b", v ->
              { acc with bold = Some (bool_of_xsd_boolean v) }
            | "charset", v ->
              { acc with charset = Some (Int32.of_string v) }
            | "condense", v ->
              { acc with condense = Some (bool_of_xsd_boolean v) }
            | "extend", v ->
              { acc with extend = Some (bool_of_xsd_boolean v) }
            | "family", v ->
              { acc with font_family = Some (Int32.of_string v) }
            | "i", v ->
              { acc with italic = Some (bool_of_xsd_boolean v) }
            | "outline", v ->
              { acc with outline = Some (bool_of_xsd_boolean v) }
            | "rFont", v ->
              { acc with font = Some v }
            | "shadow", v ->
              { acc with shadow = Some (bool_of_xsd_boolean v) }
            | "stike", v ->
              { acc with strikethrough = Some (bool_of_xsd_boolean v) }
            | "sz", v ->
              { acc with font_size = Some (Float.of_string v) }
            | "u", v  ->
              { acc with underline = Some (Underline.of_string v) }
            | "vertAlign", v ->
              { acc with vertical_align = Some (Vertical_align.of_string v) }
            | _ -> acc)
        | Element ("t", _, children) ->
          { acc with text = expect_pcdata children }
        | _ -> acc)
      |> Option.some
    | Element ("t", _, children) ->
      Some { empty with text = expect_pcdata children }
    | _ -> None
end

module String_item = struct
  type t = Rich_text.t list
      [@@deriving sexp_of]

  let to_string t =
    List.map t ~f:Rich_text.text
    |> String.concat ~sep:""

  let of_xml = function
    | Xml.Element ("si", _, children) ->
      List.filter_map children ~f:Rich_text.of_xml
      |> Option.some
    | _ -> None
end

type t = String_item.t list
    [@@deriving sexp_of]

let to_string_array t =
  List.map t ~f:String_item.to_string
  |> List.to_array

let of_xml =
  expect_element "sst" (fun _ ->
    List.filter_map ~f:String_item.of_xml)
