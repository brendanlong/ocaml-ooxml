build:
	@jbuilder build @install --dev

clean:
	@jbuilder clean

coverage: clean
	@BISECT_ENABLE=YES jbuilder runtest --dev
	@bisect-ppx-report -I _build/default/ -html _coverage/ \
	  `find . -name 'bisect*.out'`

test:
	@jbuilder runtest --dev

.PHONY: build clean coverage test
