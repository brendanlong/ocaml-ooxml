open Base
open Base.Printf
open Stdint
open Spreadsheetml

module Value = struct
  type t =
    | Date of Ptime.date
    | Datetime of Ptime.t
    | Number of float
    | String of string
    | Time of Ptime.time

  let to_string = function
    | Date (y, m, d) -> sprintf "%d-%d-%d" y m d
    | Datetime t -> Ptime.to_rfc3339 t
    | Number f ->
      Float.to_string f
      |> String.rstrip ~drop:(function '.' -> true | _ -> false)
    | String s -> s
    | Time ((h, m, s), _tz) -> sprintf "%d:%d:%d" h m s

  let sexp_of_t = function
    | Number n -> Float.sexp_of_t n
    | String s -> String.sexp_of_t s
    | _ as t -> to_string t |> String.sexp_of_t

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
    |> Map.of_alist_exn (module Int)

  let classify_format str =
    (* FIXME: This is really slow.  We should really use Menhir for this. *)
    let str = String.lowercase str in
    let remove_strings = Str.regexp "\\[[^\\[]*\\]\\|\"[^\"]*\"" in
    let str = Str.global_replace remove_strings "" str in
    if String.(str = "general" || str = "") then
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
                    || String.is_substring str ~substring:"am/pm"
                    || String.is_substring str ~substring:"a/p" in
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
        if String.(a = b) then Some a
        else if String.(a = "general" && b = "General") then Some a
        else
          failwithf "Got format string with ID %d, \"%s\", but there is a \
                     built-in format string with the same ID, \"%s\""
            key b a ())
    in
    let str = Worksheet.Cell.to_string ~shared_strings cell in
    match data_type with
    | Number when String.(str <> "") ->
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
        let xlsx_epoch = Option.value_exn ~here:[%here]
                           (Ptime.(of_date (1899, 12, 30))) in
        let date_of_float n =
          Float.iround_exn ~dir:`Down n
          |> (fun days -> Ptime.Span.of_d_ps (days, Int64.zero))
          |> Option.value_exn ~here:[%here]
          |> Ptime.add_span xlsx_epoch
          |> Option.value_exn ~here:[%here]
          |> Ptime.to_date
        in
        let time_of_float n =
          Float.((modf n |> Parts.fractional) * 24. * 60. * 60.)
          |> Ptime.Span.of_float_s
          |> Option.value_exn ~here:[%here]
          |> Ptime.of_span
          |> Option.value_exn ~here:[%here]
          |> Ptime.to_date_time
          |> snd
        in
        let n = Float.of_string str in
        function
        | `Number -> Number n
        | `Date -> Date (date_of_float n)
        | `Datetime ->
          let date = date_of_float n in
          let time = time_of_float n in
          Datetime (Ptime.of_date_time (date, time)
                    |> Option.value_exn ~here:[%here])
        | `String -> String str
        | `Time ->
          Time (time_of_float n))
    | _ -> String str

  let is_empty = function
    | String "" -> true
    | _ -> false
end

type sheet =
  { name : string
  ; rows : Value.t list list }
[@@deriving fields, sexp_of]

type t = sheet list [@@deriving sexp_of]

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
      |> Map.of_alist_exn (module Int)
    in
    let rel_map =
      zip_entry_to_xml zip "xl/_rels/workbook.xml.rels"
      |> Open_packaging.Relationships.of_xml
      |> List.map ~f:(fun { Open_packaging.Relationship.id ; target ; _ } ->
        id, target)
      |> Map.of_alist_exn (module String)
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
          |> List.max_elt ~compare:Int.compare
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
              |> Map.of_alist_exn (module Int)
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
          |> Map.of_alist_exn (module Int)
        in
        let n =
          Map.keys row_map
          |> List.max_elt ~compare:Int.compare
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
