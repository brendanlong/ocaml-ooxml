open Core_kernel
open Stdint

let sexp_of_uint32 t =
  Uint32.to_string t
  |> Sexp.of_string

let expect_element expect_name f =
  let open Xml in
  function
  | Element (name, attrs, children) when name = expect_name ->
    f attrs children
  | Element (name, _, _) ->
    failwithf "Expected <%s> element but saw <%s>" expect_name name ()
  | PCData str ->
    failwithf "Expected <%s> element but saw '%s'" expect_name str ()

let expect_pcdata children =
  List.find_map children ~f:(
    let open Xml in
    function
    | PCData str -> Some str
    | Element (name, _, _) -> failwithf "Expected PCData but got <%s>" name ())
  |> Option.value ~default:""

let require_attribute element name =
  function
  | Some v -> v
  | None -> failwithf "<%s> missing required attribute '%s'" element name ()

let bool_of_xsd_boolean =
  (* See http://books.xmlschemata.org/relaxng/ch19-77025.html *)
  function
  | "true" | "1" -> true
  | "false" | "0" -> false
  | str -> failwithf "Expected xsd:boolean but got '%s'" str ()
