build:
	@jbuilder build @install

clean:
	@jbuilder clean

doc:
	@jbuilder build @doc

test:
	@jbuilder runtest

.PHONY: build clean doc test
