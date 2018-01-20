(* See ECMA 376 Part 2 - "Open Packaging Conventions" *)
open Core_kernel
open Stdint
open Utils

module Relationship = struct
  type target_mode =
    | Internal
    | External
        [@@deriving sexp_of]

  let target_mode_of_string = function
    | "Internal" -> Internal
    | "External" -> External
    | str -> failwithf "Expected ST_TargetMode but got '%s'" str ()

  (* 8.3.3.2 Relationship Element *)
  type t =
    { target_mode : target_mode
    ; target : string (* xsd:anyURI *)
    ; type_ : string (* xsd:anyURI *)
    ; id : string (* xsd:ID *) }
      [@@deriving fields, sexp_of]

  let of_xml = function
    | Xml.Element ("Relationship", attrs,  _) ->
      let target_mode = ref Internal in
      let target = ref None in
      let type_ = ref None in
      let id = ref None in
      List.iter attrs ~f:(function
      | "TargetMode", v -> target_mode := target_mode_of_string v
      | "Target", v -> target := Some v
      | "Type", v -> type_ := Some v
      | "Id", v -> id := Some v
      | _ -> ());
      let require_attribute = require_attribute "Relationship" in
      let target = require_attribute "Target" !target in
      let type_ = require_attribute "Type" !type_ in
      let id = require_attribute "Id" !id in
      Some { target_mode = !target_mode ; target ; type_ ; id }
    | _ -> None
end

module Relationships = struct
  (* 8.3.3.1 Relationships Element *)
  type t = Relationship.t list
      [@@deriving sexp_of]

  let of_xml =
    expect_element "Relationships" (fun _ children ->
      List.filter_map children ~f:Relationship.of_xml)
end
