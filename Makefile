UV ?= uv

# the first target is the default, so just run help
default: help

test: ## Runs the unit tests
	$(UV) run python -m unittest discover -s tests -vv --locals

coverage: ## Runs unit tests and measures coverage
	$(UV) run coverage run --branch --source=reminders -m unittest discover -s tests -v
	$(UV) run coverage report -m
	$(UV) run coverage html

clean: ## Removes build artifacts
	rm -rf .coverage htmlcov/

lint: ## Check code formatting
	$(UV) run ruff check

help: ## This message
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}'
