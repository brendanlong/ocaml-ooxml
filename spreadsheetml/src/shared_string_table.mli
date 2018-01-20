module Rich_text : sig
  module Underline : sig
    type t =
      | Single
      | Double
      | Single_accounting
      | Double_accounting
      | None
        [@@deriving sexp_of]

    val of_string : string -> t
  end

  module Vertical_align : sig
    type t =
      | Baseline
      | Superscript
      | Subscript
        [@@deriving sexp_of]

    val of_string : string -> t
  end

  type t =
    { text : string
    ; bold : bool option
    ; charset : int32 option
    ; condense : bool option
    ; extend : bool option
    ; font : string option
    ; font_family : int32 option
    ; font_size : float option
    ; italic : bool option
    ; outline : bool option
    ; shadow : bool option
    ; strikethrough : bool option
    ; underline : Underline.t option
    ; vertical_align : Vertical_align.t option }
      [@@deriving fields, sexp_of]

  val empty : t

  val of_xml : Xml.xml -> t option
end

module String_item : sig
  type t = Rich_text.t list
      [@@deriving sexp_of]

  val to_string : t -> string

  val of_xml : Xml.xml -> t option
end

type t = String_item.t list
    [@@deriving sexp_of]

val to_string_array : t -> string array

val of_xml : Xml.xml -> t
