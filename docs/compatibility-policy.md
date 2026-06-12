# Compatibility Policy

This repository keeps public API compatibility as a release constraint for the `2.x` line.

## Stable surface

- Public package names stay unchanged.
- Public namespaces stay unchanged.
- Public interfaces stay source-compatible unless a release note explicitly marks a major version change.
- Existing overloads remain available; new behavior should be added through new overloads or new types.

## Allowed changes in `2.x`

- Internal refactoring.
- Dependency upgrades that do not change the public contract.
- Build and CI improvements.
- New APIs that are additive.
- Bug fixes behind the existing contract.

## Breaking changes

- Breaking changes are reserved for the next major version.
- If a public API must change, keep a compatibility wrapper in `2.x` first and mark it obsolete before removal.

## Test rule

- Any change touching public behavior must add or update behavior tests.
- CI should fail if public API drift is introduced accidentally.
