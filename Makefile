.PHONY: all check test pdf clean
all:
	npm run build
check:
	npm run typecheck
	npm run check
test:
	npm test
pdf:
	npm run build -- --pdf
clean:
	rm -rf build .cache
