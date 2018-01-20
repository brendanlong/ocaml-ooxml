build:
	@jbuilder build @install

clean:
	@jbuilder clean

coverage: clean
	@BISECT_ENABLE=YES jbuilder runtest
	@bisect-ppx-report -I _build/default/ -html _coverage/ \
	  `find . -name 'bisect*.out'`

test:
	@jbuilder runtest

.PHONY: build clean coverage test
