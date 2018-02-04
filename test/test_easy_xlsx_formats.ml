open Core_kernel
open OUnit2

(* Test that we properly classify formats from OpenOffice *)

let test_easy_xlsx_formats =
  let expect_number = -12345.6789 in
  let expect_date = Date.of_string "1866-03-12" in
  let expect_datetime =
    let time = Time.Ofday.of_string_iso8601_extended "07:42:23.039999" in
    Time.utc_mktime expect_date time in
  let expect_time = Time.Ofday.of_string_iso8601_extended "16:17:37" in
  let expect_string = "-12345.6789" in
  Easy_xlsx.read_file "test/files/formats.xlsx"
  |> List.hd_exn
  |> Easy_xlsx.rows
  |> List.tl_exn
  |> List.concat_mapi ~f:(
    fun row ->
      let row = row + 2 in
      let open Easy_xlsx.Value in
      function
      | _boolean :: number :: date :: datetime :: string :: time :: _  ->
        let wrong_type expect_type v =
          sexp_of_t v
          |> Sexp.to_string_hum
          |> sprintf "Expected %s on row %d but got %s" expect_type row
          |> assert_failure
        in
        let test type_ value f =
          if is_empty value then
            None
          else
            (sprintf "%s on row %d" type_ row
             >:: fun _ ->
               f value)
            |> Option.some
        in
        [ test "Number" number (function
          | Number n -> assert_equal ~printer:Float.to_string expect_number n
          | v -> wrong_type "Number" v)
        ; test "Date" date (function
          | Date n -> assert_equal ~printer:Date.to_string expect_date n
          | v -> wrong_type "Date" v)
        ; test "Datetime" datetime (function
        | Datetime n -> assert_equal ~printer:Time.to_string expect_datetime n
        | v -> wrong_type "Datetime" v)
        ; test "String" string (function
        | String n -> assert_equal ~printer:Fn.id expect_string n
        | v -> wrong_type "String" v)
        ; test "Time" time (function
        | Time n -> assert_equal ~printer:Time.Ofday.to_string expect_time n
        | v -> wrong_type "Time" v) ]
      | _ -> assert false)
  |> List.filter_opt
  |> List.rev

let () =
  test_easy_xlsx_formats
  |> test_list
  |> run_test_tt_main
