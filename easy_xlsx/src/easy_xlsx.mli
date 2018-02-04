(** An easy read for XLSX files. *)
open Core_kernel

module Value : sig
  type t =
    | Date of Date.t
    | Datetime of Time.t
    | Number of float
    | String of string
    | Time of Time.Ofday.t
        [@@deriving compare, sexp_of]
        (** A cell value. Note that a [Number] could be either an integer or
           a float; Excel stores them exactly the same and the only difference
           is formatting. *)

  val to_string : t -> string
  val is_empty : t -> bool
end

type sheet =
  { name : string
  ; rows : Value.t list list }
    [@@deriving compare, fields, sexp_of]
    (** One sheet from a workbook. Empty rows will be returned, and empty or
       missing columns will be returned as [""]. *)

type t = sheet list [@@deriving compare]

val read_file : string -> sheet list
  (** [read_file file_name] synchronously reads the .xlsx document at
     [file_name] and returns all of the sheets in the document. Will throw
     an exception if the file doesn't exist, can't be read for any reason,
     isn't a valid ZIP file, or isn't a valid XLSX file. *)
