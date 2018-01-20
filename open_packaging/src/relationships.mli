type t = Relationship.t list
    [@@deriving sexp_of]

val of_xml : Xml.xml -> Relationship.t list
