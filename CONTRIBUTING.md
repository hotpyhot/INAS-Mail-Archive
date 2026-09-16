# Contributing

Contributions are welcome. By submitting a contribution, you agree that it may be distributed under the Apache License 2.0 used by this project.

1. Create a branch from the current Community Edition branch.
2. Keep internal/company-specific settings, paths, domains, mail data, logs, and credentials out of commits.
3. Use dummy values such as `example@example.com`, `example.com`, and `C:\Example\...` in tests and documentation.
4. Run `python -m compileall src` and the checks in `docs/RELEASE_CHECKLIST.md`.
5. Document new dependencies and their licenses in `THIRD_PARTY_LICENSES.md`.

The internal `INAS Mail Archive Ver.1.7.x` line is a separate product line. Community Edition changes are not automatically merged into it.
