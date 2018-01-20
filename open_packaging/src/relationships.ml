open Core_kernel
open Stdint
open Utils

(* 8.3.3.1 Relationships Element *)
type t = Relationship.t list
    [@@deriving sexp_of]

let of_xml =
  expect_element "Relationships" (fun _ ->
    List.filter_map ~f:Relationship.of_xml)
