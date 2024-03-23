python_cmd ?= python3

# the first target is the default, so just run help
default: help

test: ## Runs the unit tests
	$(python_cmd) -m unittest -v 

lint: ## Check code formatting
	$(python_cmd) -m flake8

help: ## This message
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}'
