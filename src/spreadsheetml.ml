(* See ECMA 376 Part 4, Section 3 - "SpreadsheetML Reference Material" *)
open Core_kernel
open Stdint
open Utils

module Workbook = struct
  module Book_view = struct
    module Visibility = struct
      type t =
        | Hidden
        | Very_hidden
        | Visible
          [@@deriving sexp_of]

      let of_string =
        function
        | "hidden" -> Hidden
        | "veryHidden" -> Very_hidden
        | "visible" -> Visible
        | str -> failwithf "Expected ST_Visibility but got '%s'" str ()
    end

    (* 3.2.30 workbookView *)
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

    let default =
      { active_tab = Uint32.zero
      ; autofilter_date_grouping = true
      ; first_sheet = Uint32.zero
      ; minimized = false
      ; show_horizontal_scroll = true
      ; show_vertical_scroll = true
      ; tab_ratio = Uint32.of_int 600
      ; visibility = Visibility.Visible
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
              { t with visibility = Visibility.of_string v }
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

    module State = struct
      type t =
        | Hidden
        | Very_hidden
        | Visible
          [@@deriving sexp_of]

      let of_string =
        function
        | "hidden" -> Hidden
        | "veryHidden" -> Very_hidden
        | "visible" -> Visible
        | str -> failwithf "Expected ST_SheetState but got '%s'" str ()
    end

    type t =
      { id : string (* ST_RelationshipId *)
      ; name : string
      ; sheet_id : uint32
      ; state : State.t }
        [@@deriving fields, sexp_of]

    let of_xml = function
      | Xml.Element ("sheet", attrs, _) ->
        let name = ref None in
        let sheet_id = ref None in
        let state = ref State.Visible in
        let id = ref None in
        List.iter attrs ~f:(function
        | "name", v -> name := Some v
        | "sheetId", v -> sheet_id := Some (Uint32.of_string v)
        | "state", v -> state := State.of_string v
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
    expect_element "workbook" (fun _ ->
      List.fold ~init:empty ~f:(fun t ->
        let open Xml in
        function
        | Element ("bookViews", _, children) ->
          let book_views = List.filter_map children ~f:Book_view.of_xml in
          { t with book_views }
        | Element ("sheets", _, children) ->
          let sheets = List.filter_map children ~f:Sheet.of_xml in
          { t with sheets }
        | _ -> t))

end

module Shared_string_table = struct
  module Rich_text = struct
  (* Note: We represent text and rich text the same internally. Unformatted
     text is just rich text with empty formatting options *)

    module Underline = struct
      (* http://www.datypic.com/sc/ooxml/t-ssml_ST_UnderlineValues.html *)
      type t =
        | Single
        | Double
        | Single_accounting
        | Double_accounting
        | None
          [@@deriving sexp_of]

      let of_string = function
        | "single" -> Single
        | "double" -> Double
        | "singleAccounting" -> Single_accounting
        | "doubleAccounting" -> Double_accounting
        | "none" -> None
        | str -> failwithf "Expected ST_Underline but got '%s'" str ()
    end

    module Vertical_align = struct
      (* http://www.datypic.com/sc/ooxml/t-ssml_ST_VerticalAlignRun.html *)
      type t =
        | Baseline
        | Superscript
        | Subscript
          [@@deriving sexp_of]

      let of_string = function
        | "baseline" -> Baseline
        | "superscript" -> Superscript
        | "subscript" -> Subscript
        | str -> failwithf "Expected ST_VerticalAlign but got '%s'" str ()
    end

    type t =
      { text : string
      ; bold : bool option
      ; charset : int32 option
      (* ; color : CT_Color *)
      ; condense : bool option
      ; extend : bool option
      ; font : string option
      ; font_family : int32 option
      ; font_size : float option (* CT_FontSize *)
      (* ; font_scheme : CT_FontScheme *)
      ; italic : bool option
      ; outline : bool option
      ; shadow : bool option
      ; strikethrough : bool option
      ; underline : Underline.t option
      ; vertical_align : Vertical_align.t option }
        [@@deriving fields, sexp_of]

    let empty =
      { text = ""
      ; bold = None
      ; charset = None
      ; condense = None
      ; extend = None
      ; font = None
      ; font_family = None
      ; font_size = None
      ; italic = None
      ; outline = None
      ; shadow = None
      ; strikethrough = None
      ; underline = None
      ; vertical_align = None }

    let of_xml =
      let open Xml in
      function
      | Element ("r", _, children) ->
        List.fold children ~init:empty ~f:(fun acc -> function
          | Element ("rPr", attrs, _) ->
            List.fold attrs ~init:acc ~f:(fun acc -> function
              | "b", v ->
                { acc with bold = Some (bool_of_xsd_boolean v) }
              | "charset", v ->
                { acc with charset = Some (Int32.of_string v) }
              | "condense", v ->
                { acc with condense = Some (bool_of_xsd_boolean v) }
              | "extend", v ->
                { acc with extend = Some (bool_of_xsd_boolean v) }
              | "family", v ->
                { acc with font_family = Some (Int32.of_string v) }
              | "i", v ->
                { acc with italic = Some (bool_of_xsd_boolean v) }
              | "outline", v ->
                { acc with outline = Some (bool_of_xsd_boolean v) }
              | "rFont", v ->
                { acc with font = Some v }
              | "shadow", v ->
                { acc with shadow = Some (bool_of_xsd_boolean v) }
              | "stike", v ->
                { acc with strikethrough = Some (bool_of_xsd_boolean v) }
              | "sz", v ->
                { acc with font_size = Some (Float.of_string v) }
              | "u", v  ->
                { acc with underline = Some (Underline.of_string v) }
              | "vertAlign", v ->
                { acc with vertical_align = Some (Vertical_align.of_string v) }
              | _ -> acc)
          | Element ("t", _, children) ->
            { acc with text = expect_pcdata children }
          | _ -> acc)
        |> Option.some
      | Element ("t", _, children) ->
        Some { empty with text = expect_pcdata children }
      | _ -> None
  end

  module String_item = struct
    type t = Rich_text.t list
        [@@deriving sexp_of]

    let to_string t =
      List.map t ~f:Rich_text.text
      |> String.concat ~sep:""

    let of_xml = function
      | Xml.Element ("si", _, children) ->
        List.filter_map children ~f:Rich_text.of_xml
        |> Option.some
      | _ -> None
  end

  type t = String_item.t list
      [@@deriving sexp_of]

  let to_string_array t =
    List.map t ~f:String_item.to_string
    |> List.to_array

  let of_xml =
    expect_element "sst" (fun _ ->
      List.filter_map ~f:String_item.of_xml)
end
