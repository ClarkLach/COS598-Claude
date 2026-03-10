# CLAUDE.md — AI Assistant Guide for COS598-Claude

This file provides guidance for AI assistants (Claude Code and similar tools) working in this repository. Update this file as the project evolves.

---

## Repository Overview

**COS598-Claude** is the project repository for COS598, a graduate-level course. This repository uses Claude (Anthropic's AI) as a core component or subject of study.

> **Note:** This CLAUDE.md was created at repository initialization. Update the sections below as the codebase grows.

---

## Repository Structure

```
COS598-Claude/
├── CLAUDE.md           # This file — AI assistant instructions
├── README.md           # Project overview for humans
├── src/                # Main source code (add as project develops)
├── tests/              # Test suite
├── docs/               # Documentation
└── scripts/            # Utility scripts
```

Update this tree as directories are added.

---

## Development Workflow

### Branching

- **Main branch**: `main` (protected)
- **Feature branches**: `feature/<short-description>`
- **Claude-managed branches**: `claude/<session-id>` (auto-created by Claude Code)
- Always branch off `main` for new work. Never push directly to `main`.

### Commits

- Write clear, imperative commit messages: `Add feature X`, `Fix bug in Y`, `Refactor Z`
- Keep commits focused and atomic — one logical change per commit
- Reference issue numbers where relevant: `Fix login crash (#42)`

### Pull Requests

- Open a PR for every change to `main`
- Include a description of what changed and why
- Ensure tests pass before requesting review

---

## Running the Project

> **Fill in** the actual commands once the project is scaffolded.

```bash
# Install dependencies
# e.g., pip install -r requirements.txt
#     or npm install

# Run the application
# e.g., python src/main.py
#     or npm start

# Run tests
# e.g., pytest
#     or npm test
```

---

## Testing

- Write tests for all new functionality
- Place tests in the `tests/` directory, mirroring the structure of `src/`
- Tests should be deterministic — avoid relying on external services or real API calls in unit tests; use mocks/stubs
- Run the full test suite before pushing

---

## Working with the Claude API

If this project uses the Anthropic Claude API, follow these conventions:

- **API keys** must never be committed. Use environment variables (`ANTHROPIC_API_KEY`)
- Store secrets in a `.env` file (gitignored) or your system's secret manager
- Default to the latest capable model: `claude-sonnet-4-6` for general tasks, `claude-opus-4-6` for complex reasoning
- Always handle API errors gracefully (rate limits, network errors, etc.)
- Log token usage during development to avoid unexpected costs

### Example SDK usage (Python)

```python
import anthropic

client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env

message = client.messages.create(
    model="claude-sonnet-4-6",
    max_tokens=1024,
    messages=[{"role": "user", "content": "Hello, Claude!"}],
)
print(message.content[0].text)
```

---

## Code Conventions

### General

- Prefer clarity over cleverness
- Keep functions small and single-purpose
- Avoid premature abstraction — don't generalize until there are at least two concrete use cases
- Delete dead code rather than commenting it out

### Python (if applicable)

- Python 3.11+
- Follow [PEP 8](https://peps.python.org/pep-0008/); use `black` for formatting, `ruff` for linting
- Use type hints for all function signatures
- Use `pathlib.Path` instead of `os.path`

### JavaScript/TypeScript (if applicable)

- Use TypeScript with strict mode enabled
- Use `prettier` for formatting, `eslint` for linting
- Prefer `const` over `let`; avoid `var`
- Use named exports over default exports for better refactoring support

---

## Security

- **Never** commit credentials, API keys, tokens, or secrets
- Validate all external input at system boundaries
- Keep dependencies up to date; check for known vulnerabilities
- Do not log sensitive user data

---

## AI Assistant Instructions

When working in this repo, Claude Code should:

1. **Read before editing** — always read relevant files before modifying them
2. **Stay focused** — only change what was requested; avoid scope creep
3. **Keep tests green** — run tests after changes; fix failures before pushing
4. **Use the right branch** — develop on the designated `claude/` branch, never push to `main` directly
5. **Prefer editing over creating** — modify existing files rather than adding new ones when possible
6. **Ask when uncertain** — if the task is ambiguous, ask for clarification rather than guessing
7. **Follow existing patterns** — match the style and conventions already present in the codebase
8. **No security regressions** — never introduce SQL injection, XSS, command injection, or other OWASP vulnerabilities

---

## Updating This File

Keep this file current as the project evolves:
- Add new directories to the structure diagram when they're created
- Update the "Running the Project" section with actual commands
- Document any non-obvious conventions that emerge
- Remove placeholder text as it's replaced with real content
