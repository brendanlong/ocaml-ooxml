(** A parser for XLSX files. *)

type sheet =
  { name : string
  ; rows : string list list }
    [@@deriving compare, sexp]
    (** One sheet from a workbook. Empty rows will be returned, and empty or
       missing columns will be returned as [""]. *)

type t = sheet list [@@deriving compare, sexp]

val read_file : string -> sheet list
  (** [read_file file_name] synchronously reads the .xlsx document at
     [file_name] and returns all of the sheets in the document. Will throw
     an exception if the file doesn't exist, can't be read for any reason,
     isn't a valid ZIP file, or isn't a valid XLSX file. *)
