# Maintaining the public distribution

The private upstream project owns implementation, tests, version files, public documentation templates, and GitHub workflow templates.
This repository is a downstream component. It does not replace the upstream CLI.
The exporter selects files explicitly and records their SHA-256 hashes in `upstream-manifest.json`.
Its fingerprint identifies the exported content. It is not a signature.

Port fixes upstream, verify them there, then export again.
The exporter refuses to overwrite a downstream file that differs from its last exported hash and the new upstream content.
It deletes only files recorded in the previous manifest that the export no longer owns.
Files outside that manifest are preserved.
Never hand-edit the manifest to hide a conflict.

## Prepare a release

1. Set the same stable version in upstream `package.json` and `src/version.ts`.
2. Run upstream tests, typechecking, and docs generation.
3. Apply the export to this checkout, inspect the diff, and run its checks.
4. Get approval for the release commit, tag, and push.
5. Commit the export, create `vMAJOR.MINOR.PATCH`, and push the branch and tag atomically.

The tag workflow verifies the tag matches both version files, runs tests and typechecking, cross-builds four binaries on Linux, and publishes checksums.
It also attaches the upstream manifest and installer scripts.
Release notes come from GitHub. Review them before announcing the release.
The workflow needs repository Contents write permission only for publication.

Public binaries carry the `github` distribution marker and update from this repository's latest stable release.
Upstream builds carry the `source` marker and refuse binary replacement by the public updater.
Update checks never require Teams credentials.

The checksum file and binaries come from the same GitHub release. Checksums detect corruption but do not provide independent publisher verification.
Native CI smoke checks exercise startup on Linux, macOS, and Windows. The release workflow smoke-tests its Linux binary.
Release builds use Bun cross-compilation to avoid waiting for platform runner capacity.
The Release workflow can also rebuild an existing tag through its manual tag input.
Human browser sign-in and screenshot checks remain separate.
