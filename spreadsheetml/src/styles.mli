open Stdint

module Format : sig
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

  val default : t

  val of_xml : Xml.xml -> t option
end

module Number_format : sig
  type t =
    { id : uint32
    ; format : string }
      [@@deriving fields, sexp_of]

  val of_xml : Xml.xml -> t option
end

type t =
  { cell_formats : Format.t list
  ; formatting_records : Format.t list
  ; number_formats : Number_format.t list }
    [@@deriving fields, sexp_of]

val empty : t

val of_xml : Xml.xml -> t
