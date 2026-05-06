# Cross-reference-Agent

## Cursor Cloud specific instructions

This repository is currently empty (initial commit with only a README.md). There are no services, dependencies, build systems, or tests to run.

### When code is added

Once application code is added to this repository, future agents should:

1. Check for dependency manifest files (`package.json`, `requirements.txt`, `pyproject.toml`, etc.) and install accordingly.
2. Look for lint/test/build scripts in the manifest or a `Makefile`.
3. Update the VM environment update script via `SetupVmEnvironment` to include the appropriate dependency install command.

### Current state

- **Language/Framework**: None yet
- **Services**: None
- **Database**: None
- **External dependencies**: None
