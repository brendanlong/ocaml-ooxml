build:
	@dune build @install

clean:
	@dune clean

coverage: clean
	@BISECT_ENABLE=YES dune runtest --force
	@bisect-ppx-report html

test:
	@dune runtest --force

setup:
	@opam pin add -y -n --kind path open_packaging .
	@opam pin add -y -n --kind path spreadsheetml .
	@opam pin add -y -n --kind path easy_xlsx .
	@opam install --deps-only -y -t easy_xlsx open_packaging spreadsheetml

.PHONY: build clean coverage setup test

