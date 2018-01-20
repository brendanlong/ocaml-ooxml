module Target_mode : sig
  type t =
    | Internal
    | External
        [@@deriving sexp_of]

  val of_string : string -> t
end

type t =
  { target_mode : Target_mode.t
  ; target : string (** xsd:anyURI *)
  ; type_ : string (** xsd:anyURI *)
  ; id : string (** xsd:ID *) }
    [@@deriving fields, sexp_of]

val of_xml : Xml.xml -> t option
