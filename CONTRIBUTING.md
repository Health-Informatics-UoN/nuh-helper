# Contributing to nuh-helper

We welcome contributions. All submissions will be reviewed by the repository’s **code owners** before merge. Please follow the guidelines below so reviews can go smoothly.

## Development setup

This project uses **uv** for dependency management. Install dev dependencies (including pytest and pre-commit) with:

```bash
uv sync --group dev
```

## Pre-commit hooks (required)

You **must** use pre-commit hooks for this repo. They enforce formatting and basic checks before commit.

1. **Install the git hooks** (once per clone):

   ```bash
   uv run pre-commit install
   ```

2. **Run manually** (optional; they also run on `git commit`):

   ```bash
   uv run pre-commit run --all-files
   ```

Hooks include: trailing-whitespace, end-of-file-fixer, YAML checks, large-file checks, **Ruff** (lint + fix), and **Ruff format**. Fix any reported issues before pushing.

## Tests

Tests are in the `tests/` directory and use **pytest**.

- **Run all tests:**

  ```bash
  uv run pytest
  ```

CI runs the full test suite; ensure `uv run pytest` passes locally before opening a PR.

## Commits and PRs

- Use **Angular-style semantic commit messages** (e.g. `feat: add X`, `fix: handle Y`, `docs: update Z`). CI checks this.
- Open a PR against the default branch. A code owner will review and approve before merge.

Thank you for contributing.
