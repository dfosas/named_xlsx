# Intended for local development only

SHELL := /bin/bash
.DEFAULT_GOAL := help
.PHONY: build coverage docs-build docs-deploy docs-serve help install lint lock sync test

dir_package="named_xlsx"
dir_tests="tests"


install: sync  ## Install project and default dev groups into .venv

sync:  ## Sync local environment from pyproject.toml and uv.lock
	uv sync --extra xlsx

lock:  ## Refresh uv.lock after dependency or metadata changes
	uv lock

build:  ## Build source and wheel distributions
	uv build

docs-serve:  ## Preview documentation locally
	uv run zensical serve

docs-build:  ## Build documentation site
	uv run zensical build

docs-deploy:  ## Build documentation and publish to gh-pages
	uv run zensical build
	touch site/.nojekyll
	uv run ghp-import -n -p -f site

lint:  ## Lint and static-check
	uv run black --check --diff --color --workers 1 $(dir_package) $(dir_tests)
	uv run python -m flake8 $(dir_package) $(dir_tests)
	uv run python -m mypy $(dir_package) $(dir_tests)
	uv run python -m pylint $(dir_package)

test:  ## Test
	uv run coverage erase
	uv run pytest
	uv run coverage report

help: ## Show help message
	@IFS=$$'\n' ; \
	help_lines=(`fgrep -h "##" $(MAKEFILE_LIST) | fgrep -v fgrep | sed -e 's/\\$$//' | sed -e 's/##/:/'`); \
	printf "%s\n\n" "Usage: make [task]"; \
	printf "%-15s %s\n" "Task"            "Help"                          ; \
	printf "%-15s %s\n" "---------------" "-----------------------------" ; \
	for help_line in $${help_lines[@]}; do \
		IFS=$$':' ; \
		help_split=($$help_line) ; \
		help_command=`echo $${help_split[0]} | sed -e 's/^ *//' -e 's/ *$$//'` ; \
		help_info=`echo $${help_split[2]} | sed -e 's/^ *//' -e 's/ *$$//'` ; \
		printf '\033[36m'; \
		printf "%-15s %s" $$help_command ; \
		printf '\033[0m'; \
		printf "%s\n" $$help_info; \
	done
