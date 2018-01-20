(* See ECMA 376 Part 4, Section 3 - "SpreadsheetML Reference Material" *)
open Core_kernel
open Stdint

let sexp_of_uint32 t =
  Uint32.to_string t
  |> Sexp.of_string

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

module Workbook = struct
  module Book_view = struct
    type visibility =
      | Hidden
      | Very_hidden
      | Visible
        [@@deriving sexp_of]

    let visibility_of_string =
      function
      | "hidden" -> Hidden
      | "veryHidden" -> Very_hidden
      | "visible" -> Visible
      | str -> failwithf "Expected ST_Visibility but got '%s'" str ()

    (* 3.2.30 workbookView *)
    type t =
      { active_tab : uint32
      ; autofilter_date_grouping : bool
      ; first_sheet : uint32
      ; minimized : bool
      ; show_horizontal_scroll : bool
      ; show_vertical_scroll : bool
      ; tab_ratio : uint32
      ; visibility : visibility
      ; window_height : uint32 option
      ; window_width : uint32 option
      ; x_window : int32 option
      ; y_window : int32 option }
        [@@deriving fields, sexp_of]

    let default =
      { active_tab = Uint32.zero
      ; autofilter_date_grouping = true
      ; first_sheet = Uint32.zero
      ; minimized = false
      ; show_horizontal_scroll = true
      ; show_vertical_scroll = true
      ; tab_ratio = Uint32.of_int 600
      ; visibility = Visible
      ; window_height = None
      ; window_width = None
      ; x_window = None
      ; y_window = None }

    let of_xml = function
      | Xml.Element ("bookView", attrs, _) ->
        List.fold attrs ~init:default ~f:(fun t ->
            function
            | "activeTab", v ->
              { t with active_tab = Uint32.of_string v }
            | "autoFilterDateGrouping", v ->
              { t with autofilter_date_grouping = bool_of_xsd_boolean v }
            | "firstSheet", v ->
              { t with first_sheet = Uint32.of_string v }
            | "minimized", v ->
              { t with minimized = bool_of_xsd_boolean v }
            | "showHorizontalScroll", v ->
              { t with show_horizontal_scroll = bool_of_xsd_boolean v }
            | "showVerticalScroll", v ->
              { t with show_vertical_scroll = bool_of_xsd_boolean v }
            | "tabRatio", v ->
              { t with tab_ratio = Uint32.of_string v }
            | "visibility", v ->
              { t with visibility = visibility_of_string v }
            | "windowHeight", v ->
              { t with window_height = Some (Uint32.of_string v) }
            | "windowWidth", v ->
              { t with window_width = Some (Uint32.of_string v) }
            | "xWindow", v ->
              { t with x_window = Some (Int32.of_string v) }
            | "yWindow", v ->
              { t with y_window = Some (Int32.of_string v) }
            | _ -> t)
        |> Option.some
      | _ -> None
  end

  module Sheet = struct
    (* 3.2.19 sheet (Sheet information) *)
    type state =
      | Hidden
      | Very_hidden
      | Visible
        [@@deriving sexp_of]

    let state_of_string =
      function
      | "hidden" -> Hidden
      | "veryHidden" -> Very_hidden
      | "visible" -> Visible
      | str -> failwithf "Expected ST_SheetState but got '%s'" str ()

    type t =
      { id : string (* ST_RelationshipId *)
      ; name : string
      ; sheet_id : uint32
      ; state : state }
        [@@deriving fields, sexp_of]

    let of_xml = function
      | Xml.Element ("sheet", attrs, _) ->
        let name = ref None in
        let sheet_id = ref None in
        let state = ref Visible in
        let id = ref None in
        List.iter attrs ~f:(function
        | "name", v -> name := Some v
        | "sheetId", v -> sheet_id := Some (Uint32.of_string v)
        | "state", v -> state := state_of_string v
        | "r:id", v -> id := Some v
        | _ -> ());
        let name = require_attribute "sheet" "name" !name in
        let sheet_id = require_attribute "sheet" "sheetId" !sheet_id in
        let id = require_attribute "sheet" "r:id" !id in
        Some { name ; sheet_id ; state = !state ; id }
      | _ -> None
  end

  type t =
    { book_views : Book_view.t list
    (*
    ; calculation_properties : Calculation_property.t list
    ; custom_workbook_views : Custom_workbook_view.t list
    ; defined_names : Defined_name.t list
    ; external_references : External_reference.t list
    ( * TODO: extLst - Future Feature Storage Area * )
    ; file_recovery_properties : File_recovery_property.t list
    ; file_sharing : File_sharing.t list
    ; file_version : File_version.t list
    ; function_groups : Function_group.t list
    ; embedded_object_sizes : Embedded_object_size.t list
    ; pivot_caches : Pivot_cache.t list *)
    ; sheets : Sheet.t list (* 3.2.20 sheets *)
    (*
    ; smart_tag_properties : Smart_tag_property.t list
    ; smart_tag_types : Smart_tag_type.t list
    ; web_publishing_properties : Web_publishing_properties.t list
    ; web_publish_objects : Web_publish_object.t list
    ; workbook_properties : Workbook_property.t list
    ; workbook_protection : Workbook_protection.t list *) }
      [@@deriving fields, sexp_of]

  let empty =
    { book_views = []
    ; sheets = [] }

  let of_xml =
    let open Xml in
    function
    | Element ("workbook", attrs, children) ->
      List.fold children ~init:empty ~f:(fun t ->
          function
          | Element ("bookViews", _, children) ->
            let book_views = List.filter_map children ~f:Book_view.of_xml in
            { t with book_views }
          | Element ("sheets", _, children) ->
            let sheets = List.filter_map children ~f:Sheet.of_xml in
            { t with sheets }
          | _ -> t)
    | Element (name, _, _) ->
      failwithf "Expected <workbook> element but saw <%s>" name ()
    | PCData str ->
      failwithf "Expected <workbook> element but saw '%s'" str ()

end
