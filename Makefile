UV_RUN ?= uv run

# the first target is the default, so just run help
default: help
	
###########
##@ General
clean: ## Removes build artifacts
	rm -rf .coverage htmlcov/

lint: ## Check code formatting
	$(UV_RUN) ruff check

help: ## This message
	@echo "===================="
	@echo " Available Commands"
	@echo "===================="
	@awk 'BEGIN {FS = ":.*##"; printf "\nUsage:\n  make \033[36m\033[0m\n"} /^[$$()% a-zA-Z_-]+:.*?##/ { printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2 } /^##@/ { printf "\n\033[1m%s\033[0m\n", substr($$0, 5) } ' $(MAKEFILE_LIST)

###########
##@ Test
test: ## Runs the unit tests
	$(UV_RUN) python -m unittest discover -s tests -vv --locals

coverage: ## Runs unit tests and measures coverage
	$(UV_RUN) coverage run --branch --source=reminders -m unittest discover -s tests -v
	$(UV_RUN) coverage report -m
	$(UV_RUN) coverage html

