build:
	@jbuilder build @install

clean:
	@jbuilder clean

test:
	@jbuilder runtest

.PHONY: build clean test
