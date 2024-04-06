python_cmd ?= python3
coverage_cmd ?= $(python_cmd) -m coverage

# the first target is the default, so just run help
default: help

test: ## Runs the unit tests
	$(python_cmd) -m unittest -v 

coverage: ## Runs unit tests and measures coverage
	$(coverage_cmd) run --branch --source=reminders -m unittest -v
	$(coverage_cmd) report -m
	$(coverage_cmd) html

clean: ## Removes build artifacts
	rm -rf .coverage htmlcov/

lint: ## Check code formatting
	$(python_cmd) -m flake8

help: ## This message
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}'
