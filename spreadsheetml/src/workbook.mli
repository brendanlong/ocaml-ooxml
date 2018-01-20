open Stdint

module Book_view : sig
  module Visibility : sig
    type t =
      | Hidden
      | Very_hidden
      | Visible
        [@@deriving sexp_of]

    val of_string : string -> t
  end

  type t =
    { active_tab : uint32
    ; autofilter_date_grouping : bool
    ; first_sheet : uint32
    ; minimized : bool
    ; show_horizontal_scroll : bool
    ; show_vertical_scroll : bool
    ; tab_ratio : uint32
    ; visibility : Visibility.t
    ; window_height : uint32 option
    ; window_width : uint32 option
    ; x_window : int32 option
    ; y_window : int32 option }
      [@@deriving fields, sexp_of]

  val default : t

  val of_xml : Xml.xml -> t option
end

module Sheet : sig

  module State : sig
    type t =
      | Hidden
      | Very_hidden
      | Visible
        [@@deriving sexp_of]

    val of_string : string -> t
  end

  type t =
    { id : string (** ST_RelationshipId *)
    ; name : string
    ; sheet_id : uint32
    ; state : State.t }
      [@@deriving fields, sexp_of]

  val of_xml : Xml.xml -> t option
end

type t =
  { book_views : Book_view.t list
  ; sheets : Sheet.t list (** 3.2.20 sheets *) }
    [@@deriving fields, sexp_of]

val empty : t

val of_xml : Xml.xml -> t
