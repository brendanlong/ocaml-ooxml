open Base
open Base.Printf
open OUnit2

(* Test that we properly classify formats from OpenOffice *)

let test_easy_xlsx_formats =
  let open Easy_xlsx.Value in
  let expect_number = Number (-12345.6789) in
  let expect_date = Date (1866, 03, 12) in
  let expect_datetime =
    let date = 1866, 03, 12 in
    let time = (7, 42, 23), 0 in
    Datetime (Ptime.of_date_time (date, time) |> Option.value_exn) in
  let expect_time = Time ((16, 17, 37), 0) in
  let expect_string = String "-12345.6789" in
  Easy_xlsx.read_file "easy_xlsx/test/files/formats.xlsx"
  |> List.hd_exn
  |> Easy_xlsx.rows
  |> List.tl_exn
  |> List.concat_mapi ~f:(
    fun row ->
      let row = row + 2 in
      let open Easy_xlsx.Value in
      function
      | _boolean :: number :: date :: datetime :: string :: time :: _  ->
        let test type_ value expected =
          if is_empty value then
            None
          else
            (sprintf "%s on row %d" type_ row
             >:: fun _ ->
               assert_equal ~printer:Easy_xlsx.Value.to_string expected value)
            |> Option.some
        in
        [ test "Number" number expect_number
        ; test "Date" date expect_date
        ; test "Datetime" datetime expect_datetime
        ; test "String" string expect_string
        ; test "Time" time expect_time ]
      | _ -> assert false)
  |> List.filter_opt
  |> List.rev

let () =
  test_easy_xlsx_formats
  |> test_list
  |> run_test_tt_main
