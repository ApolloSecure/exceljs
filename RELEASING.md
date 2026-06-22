# Releasing ApolloSecure ExcelJS

This fork keeps upstream ExcelJS SemVer as the base version and adds an Apollo-specific release suffix for fork-only changes.

## Tag Format

Use annotated tags in this format:

`v<upstream-version>-apollo.<n>`

Examples:

- `v4.4.0-apollo.1`
- `v4.4.0-apollo.2`
- `v4.4.1-apollo.1`

Rules:

- `<upstream-version>` is the ExcelJS version this fork is based on.
- `apollo.<n>` increments for fork-only releases on the same upstream base.
- When the fork rebases or upgrades to a new upstream version, reset the Apollo counter to `.1`.

## Commit Convention

Use conventional commits for release-worthy changes.

Examples:

- `fix(table): ignore orphaned auto-filter columns`
- `docs(release): document Apollo release workflow`
- `feat(xlsx): add support for ...`

## Release Steps

1. Make sure the branch is clean and tests for the change have passed.
2. Create a conventional commit for the release contents.
3. Create an annotated tag:

   `git tag -a v4.4.0-apollo.1 -m "release: v4.4.0-apollo.1"`

4. Push the branch and tag:

   `git push origin master`

   `git push origin v4.4.0-apollo.1`

## Choosing The Next Tag

- If the last released tag is `v4.4.0-apollo.1`, the next fork-only release is `v4.4.0-apollo.2`.
- If upstream moves to `4.4.1` and this fork adopts it, the next fork release becomes `v4.4.1-apollo.1`.

## Current Baseline

- Repository: `ApolloSecure/exceljs`
- First Apollo tag: `v4.4.0-apollo.1`
