open Core_kernel
open Stdint
open Utils

module Target_mode = struct
  type t =
    | Internal
    | External
        [@@deriving sexp_of]

  let of_string = function
    | "Internal" -> Internal
    | "External" -> External
    | str -> failwithf "Expected ST_TargetMode but got '%s'" str ()
end

(* 8.3.3.2 Relationship Element *)
type t =
  { target_mode : Target_mode.t
  ; target : string
  ; type_ : string
  ; id : string }
    [@@deriving fields, sexp_of]

let of_xml = function
  | Xml.Element ("Relationship", attrs,  _) ->
    let target_mode = ref Target_mode.Internal in
    let target = ref None in
    let type_ = ref None in
    let id = ref None in
    List.iter attrs ~f:(function
    | "TargetMode", v -> target_mode := Target_mode.of_string v
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
