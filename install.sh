#!/bin/sh
# Download and verify a public release for the current user.
set -eu
repo=LegoGuy32109/agentic-teams-cli
fail() { echo "install: $1" >&2; exit 1; }
command -v curl >/dev/null 2>&1 || fail "curl is required."
case "$(uname -s)" in Linux) os=linux ;; Darwin) os=darwin ;; *) fail "Use install.ps1 on Windows." ;; esac
case "$(uname -m)" in x86_64|amd64) arch=x64 ;; arm64|aarch64) arch=arm64 ;; *) fail "Unsupported processor." ;; esac
case "$os-$arch" in linux-x64|darwin-x64|darwin-arm64) ;; *) fail "No binary for $os-$arch." ;; esac
name=teams-cli-$os-$arch
release_url=$(curl -fsSL -o /dev/null -w '%{url_effective}' "https://github.com/$repo/releases/latest")
tag=${release_url##*/}
printf '%s\n' "$tag" | LC_ALL=C grep -Eq '^v[0-9]+\.[0-9]+\.[0-9]+$' || fail "No stable public release was found."
scratch=$(mktemp -d)
staged=
trap 'rm -rf "$scratch"; if [ -n "$staged" ]; then rm -f "$staged"; fi' EXIT HUP INT TERM
base=https://github.com/$repo/releases/download/$tag
curl -fsSL "$base/$name" -o "$scratch/$name"
curl -fsSL "$base/SHA256SUMS" -o "$scratch/SHA256SUMS"
expected=$(awk -v file="$name" '$2 == file { print $1 }' "$scratch/SHA256SUMS")
printf '%s\n' "$expected" | LC_ALL=C grep -Eq '^[a-fA-F0-9]{64}$' || fail "Missing or ambiguous checksum."
if command -v sha256sum >/dev/null 2>&1; then
  actual=$(sha256sum "$scratch/$name" | awk '{print $1}')
elif command -v shasum >/dev/null 2>&1; then
  actual=$(shasum -a 256 "$scratch/$name" | awk '{print $1}')
else
  fail "sha256sum or shasum is required."
fi
[ "$actual" = "$expected" ] || fail "Checksum mismatch. Nothing was installed."
bin_dir=${HOME:?HOME is not set}/.local/bin
mkdir -p "$bin_dir"
staged=$(mktemp "$bin_dir/.teams-cli-XXXXXX")
cp "$scratch/$name" "$staged"
chmod 755 "$staged"
if [ "$os" = darwin ]; then xattr -d com.apple.quarantine "$staged" 2>/dev/null || true; fi
mv -f "$staged" "$bin_dir/teams-cli"
echo "Installed $tag to $bin_dir/teams-cli."
case ":${PATH:-}:" in *":$bin_dir:"*) ;; *) echo 'Add to your shell profile: export PATH="$HOME/.local/bin:$PATH"' ;; esac
echo "Next: teams-cli auth login"
echo "Updates: teams-cli update"
