
# code static lint
DIRS=$(shell ls -1 -F | grep "/$$" | grep -v vendor)
.PHONY: fmt
fmt:
	@echo "==> Fixing source code with gofmt..."
	@for dir in $(DIRS) ; do `goimports -w $$dir` ; done
	@for dir in $(DIRS) ; do `gofmt -s -w $$dir` ; done


.PHONY: dep
mod:
	@echo "==> Downloading dependencies..."
	@go mod tidy
	@go mod download

.PHONY: lint
lint: fmt dep
	@echo "==> Checking golangci-lint..."
	@golangci-lint run