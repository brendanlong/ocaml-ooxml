open Core_kernel
open Stdint
open Spreadsheetml

module Value = struct
  type time = Time.t [@@deriving compare]

  let sexp_of_time t =
    Time.to_string_iso8601_basic ~zone:Time.Zone.utc t
    |> Sexp.of_string

  type t =
    | Date of Date.t
    | Datetime of time
    | Number of float
    | String of string
    | Time of Time.Ofday.t
        [@@deriving compare, sexp_of]

  let built_in_formats =
    (* FIXME: This is only the English format codes. There are 4 Asian format
       codes (Chinese Simplified, Chinese Traditional, Japanese, Korean) and it's
       not clear to me how we're supposed to pick between them. *)
    [ 0, "general"
    ; 1, "0"
    ; 2, "0.00"
    ; 3, "#,##0"
    ; 4, "#,##0,00"
    ; 9, "0%"
    ; 10, "0.00%"
    ; 11, "0.00E+00"
    ; 12, "# ?/?"
    ; 13, "# ??/??"
    ; 14, "mm-dd-yy"
    ; 15, "d-mmm-yy"
    ; 16, "d-mmm"
    ; 17, "mmm-yy"
    ; 18, "h:mm AM/PM"
    ; 19, "h:mm:ss AM/PM"
    ; 20, "h:mm"
    ; 21, "h:mm:ss"
    ; 22, "m/d/yy h:mm"
    ; 37, "#,##0 ;(#,##0)"
    ; 38, "#,##0 ;[Red](#,##0)"
    ; 39, "#,##0.00;(#,##0.00)"
    ; 40, "#,##0.00;[Red](#,##0.00)"
    ; 45, "mm:ss"
    ; 46, "[h]:mm:ss"
    ; 47, "mmss.0"
    ; 48, "##0.0E+0"
    ; 49, "@" ]
    |> Int.Map.of_alist_exn

  let classify_format str =
    (* FIXME: This is really slow.  We should really use Menhir for this. *)
    let str = String.lowercase str in
    let remove_strings = Str.regexp "\\[[^\\[]*\\]\\|\"[^\"]*\"" in
    let str = Str.global_replace remove_strings "" str in
    if str = "general" || str = "" then
      `Number
    else
      let is_string = String.contains str '@' in
      let is_date = String.contains str 'y' || String.contains str 'd'
                    (* QQ is quarter, NN is day of the week *)
                    || String.contains str 'q' || String.contains str 'n'
                    (* WW is week number *)
                    || String.contains str 'w'
                    || String.is_substring str ~substring:"mmm" in
      let is_time = String.contains str 'h' || String.contains str 's'
                    || String.is_substring str ~substring:"AM/PM"
                    || String.is_substring str ~substring:"A/P" in
      if [ is_string ; is_date || is_time ]
         |> List.filter ~f:Fn.id
         |> List.length > 1 then
        failwithf "Ambiguous format string '%s'" str ()
      else if is_string then
        `String
      else if is_date && is_time then
        `Datetime
      else if is_date then
        `Date
      else if is_time then
        `Time
      else
        `Number

  let of_cell ~styles ~formats ~shared_strings
      ({ Worksheet.Cell.style_index ; data_type ; _ } as cell) =
    let formats = Map.merge built_in_formats formats ~f:(fun ~key ->
        function
        | `Left v -> Some v
        | `Right v -> Some v
        | `Both (a, b) ->
          if a = b then Some a
          else
            failwithf "Got format string with ID %d, \"%s\", but there is a \
                       built-in format string with the same ID, \"%s\""
              key b a ())
    in
    let str = Worksheet.Cell.to_string ~shared_strings cell in
    match data_type with
    | Number when str <> "" ->
      styles.(Uint32.to_int style_index)
      |> Spreadsheetml.Styles.Format.number_format_id
      |> Option.map ~f:Uint32.to_int
      |> Option.value ~default:0
      |> (fun num_fmt_id ->
        match Map.find formats num_fmt_id with
        | Some format -> format
        | None ->
          failwithf "Cell referenced numFmtId %d but it's not listed in the \
                     XLSX file and isn't a known built-in format ID"
            num_fmt_id ())
      |> classify_format
      |> (
        let date_of_float n =
          let d = Date.create_exn ~y:1899 ~m:Month.Dec ~d:30 in
          Date.add_days d (Float.iround_exn ~dir:`Down n)
        in
        let time_of_float n =
          let us =
            let time_part = Float.(modf n |> Parts.fractional) in
            let time_abs =
              if time_part >=. 0. then time_part
              else 1. +. time_part
            in
            time_abs *. 24. *. 3600. *. 1000000.
            |> Float.to_int
          in
          Time.Ofday.create ~us ()
        in
        let n = float_of_string str in
        function
        | `Number -> Number n
        | `Date -> Date (date_of_float n)
        | `Datetime ->
          let date = date_of_float n in
          let time = time_of_float n in
          Datetime (Time.utc_mktime date time)
        | `String -> String str
        | `Time ->
          Time (time_of_float n))
    | _ -> String str

  let to_string = function
    | Date d -> Date.to_string_iso8601_basic d
    | Datetime t -> Time.to_string_iso8601_basic ~zone:Time.Zone.utc t
    | Number f ->
      Float.to_string_round_trippable f
      |> String.rstrip ~drop:(function '.' -> true | _ -> false)
    | String s -> s
    | Time t -> Time.Ofday.to_string_trimmed t

  let is_empty = function
    | String "" -> true
    | _ -> false
end

type sheet =
  { name : string
  ; rows : Value.t list list }
    [@@deriving compare, fields, sexp_of]

type t = sheet list [@@deriving compare]

let read_file filename =
  let zip_entry_to_xml zip name =
    Zip.find_entry zip name
    |> Zip.read_entry zip
    |> Xml.parse_string
  in
  let zip = Zip.open_in filename in
  Exn.protect ~f:(fun () ->
    let shared_strings =
      zip_entry_to_xml zip "xl/sharedStrings.xml"
      |> Shared_string_table.of_xml
      |> List.to_array
    in
    let sheets =
      zip_entry_to_xml zip "xl/workbook.xml"
      |> Workbook.of_xml
      |> Workbook.sheets
    in
    let stylesheet =
      zip_entry_to_xml zip "xl/styles.xml"
      |> Styles.of_xml
    in
    let styles =
      Styles.cell_formats stylesheet
      |> Array.of_list
    in
    let formats =
      Styles.number_formats stylesheet
      |> List.map ~f:(fun { Styles.Number_format.id ; format } ->
        Uint32.to_int id, format)
      |> Int.Map.of_alist_exn
    in
    let rel_map =
      zip_entry_to_xml zip "xl/_rels/workbook.xml.rels"
      |> Open_packaging.Relationships.of_xml
      |> List.map ~f:(fun { Open_packaging.Relationship.id ; target ; _ } ->
        id, target)
      |> String.Map.of_alist_exn
    in
    List.map sheets ~f:(fun { Workbook.Sheet.name ; id ; _ } ->
      let rows =
        let target = Map.find_exn rel_map id in
        let path = sprintf "xl/%s" target in
        let { Worksheet.columns ; rows } =
          Zip.find_entry zip path
          |> Zip.read_entry zip
          |> Xml.parse_string
          |> Worksheet.of_xml
        in
        let num_cols =
          columns
          |> List.map ~f:Worksheet.Column.max
          |> List.map ~f:Uint32.to_int
          |> List.max_elt ~cmp:Int.compare
          |> Option.value ~default:0
        in
        let row_map =
          rows
          |> List.map ~f:(fun { Worksheet.Row.row_index ; cells ; _ } ->
            let index =
              Option.value_exn ~here:[%here] row_index
              |> Uint32.to_int
            in
            let cell_map =
              List.map cells ~f:(fun cell ->
                Worksheet.Cell.column cell, cell)
              |> Int.Map.of_alist_exn
            in
            let cells =
              Map.max_elt cell_map
              |> Option.map ~f:(fun (max, _) ->
                List.init (max + 1) ~f:(fun i ->
                  Map.find cell_map i
                  |> Option.value ~default:Worksheet.Cell.default))
              |> Option.value ~default:[]
            in
            index - 1, cells)
          |> Int.Map.of_alist_exn
        in
        let n =
          Map.keys row_map
          |> List.max_elt ~cmp:Int.compare
          |> Option.map ~f:((+) 1)
          |> Option.value ~default:0
        in
        List.init n ~f:Fn.id
        |> List.map ~f:(fun i ->
          let row =
            Map.find row_map i
            |> Option.value ~default:[]
          in
          let missing_cols = num_cols - List.length row in
          if missing_cols > 0 then
            row @ List.init ~f:(Fn.const Worksheet.Cell.default) missing_cols
          else
            row)
        |> List.map ~f:(List.map ~f:(Value.of_cell ~styles ~formats ~shared_strings))
      in
      { name ; rows }))
    ~finally:(fun () -> Zip.close_in zip)
