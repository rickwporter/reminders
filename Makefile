UV_RUN ?= uv run

# the first target is the default, so just run help
default: help
	
###########
##@ General
all: lint coverage ## Removes build artifacts

clean: ## Removes build artifacts
	rm -rf .coverage htmlcov/ .ruff_cache/ dist/
	rm -rf `find . -name __pycache__`

help: ## This message
	@echo "===================="
	@echo " Available Commands"
	@echo "===================="
	@awk 'BEGIN {FS = ":.*##"; printf "\nUsage:\n  make \033[36m\033[0m\n"} /^[$$()% a-zA-Z_-]+:.*?##/ { printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2 } /^##@/ { printf "\n\033[1m%s\033[0m\n", substr($$0, 5) } ' $(MAKEFILE_LIST)

###########
##@ Lint
delint: ## Fix formatting issues
	$(UV_RUN) ruff check --fix

lint: ## Check code formatting
	$(UV_RUN) ruff check

###########
##@ Test
test: ## Runs the unit tests
	$(UV_RUN) python -m unittest discover -s tests -vv --locals

coverage: ## Runs unit tests and measures coverage
	$(UV_RUN) coverage run --branch --source=reminders -m unittest discover -s tests -v
	$(UV_RUN) coverage report -m
	$(UV_RUN) coverage html

