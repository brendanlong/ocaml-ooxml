#use "topfind"
#require "topkg-jbuilder"

open Topkg

let () =
  Topkg_jbuilder.describe ~name:"easy_xlsx" ();
  Topkg_jbuilder.describe ~name:"open_packaging" ();
  Topkg_jbuilder.describe ~name:"spreadsheetml" ();
