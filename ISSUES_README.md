# Code Review Issues

This directory contains a script to create GitHub issues for high-priority code review findings.

## Prerequisites

You need the GitHub CLI (`gh`) installed and authenticated:

1. **Install GitHub CLI**:
   - macOS: `brew install gh`
   - Linux: See https://github.com/cli/cli/blob/trunk/docs/install_linux.md
   - Windows: See https://cli.github.com/

2. **Authenticate**:
   ```bash
   gh auth login
   ```

## Usage

Run the script to create all 5 issues:

```bash
./create-github-issues.sh
```

The script will create the following issues:

1. **Replace Bare Exception Handlers** (Critical Priority)
   - Fixes dangerous catch-all exception handling
   - Labels: `bug`, `priority: critical`

2. **Add Proper Logging** (High Priority)
   - Replace print statements with structured logging
   - Labels: `enhancement`, `priority: high`

3. **Extract Configuration Constants** (High Priority)
   - Remove magic numbers and improve maintainability
   - Labels: `enhancement`, `priority: high`, `refactoring`

4. **Add Input Validation Function** (High Priority)
   - Comprehensive PDF and output path validation
   - Labels: `enhancement`, `priority: high`

5. **Improve Resource Management** (High Priority)
   - Fix resource leaks and ensure cleanup
   - Labels: `bug`, `priority: high`

## What the Script Does

- Checks that `gh` CLI is installed
- Verifies GitHub authentication
- Creates 5 detailed issues with:
  - Problem descriptions
  - Affected code locations
  - Recommended solutions
  - Implementation steps
  - Benefits

## Troubleshooting

**Error: "gh: command not found"**
- Install the GitHub CLI (see Prerequisites above)

**Error: "Not authenticated with GitHub CLI"**
- Run `gh auth login` to authenticate

**Error: "Permission denied"**
- Make sure the script is executable: `chmod +x create-github-issues.sh`

## After Running

View your newly created issues at:
```
https://github.com/YOUR-USERNAME/pdf-to-xls-vision/issues
```

Or use the CLI:
```bash
gh issue list
```
