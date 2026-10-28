#!/bin/sh
# Installs teams-cli for the current user on Linux or macOS.
#
# Usage: install/install.sh [DIR]
#
# DIR holds the binaries that `bun run build:all` writes. Without DIR, the script looks in ./dist and then
# in the dist folder next to this script. It downloads nothing and needs no sudo. It copies one file to
# ~/.local/bin/teams-cli.
#
# TEAMS_CLI_PLATFORM (for example linux-x64) overrides the platform that the script detects. Tests use it.
set -eu

fail() {
  echo "install: $1" >&2
  exit 1
}

if [ -n "${TEAMS_CLI_PLATFORM:-}" ]; then
  platform=$TEAMS_CLI_PLATFORM
else
  os=$(uname -s)
  machine=$(uname -m)
  case "$os" in
    Linux) os=linux ;;
    Darwin) os=darwin ;;
    *) fail "this script supports Linux and macOS. On Windows, run install/install.ps1." ;;
  esac
  case "$machine" in
    x86_64 | amd64) arch=x64 ;;
    arm64 | aarch64) arch=arm64 ;;
    *) fail "no binary for the processor type $machine." ;;
  esac
  platform=$os-$arch
fi

case "$platform" in
  linux-x64 | darwin-arm64 | darwin-x64) ;;
  *) fail "no binary for $platform. Available: linux-x64, darwin-arm64, darwin-x64." ;;
esac

name=teams-cli-$platform
here=$(cd "$(dirname "$0")" && pwd)

if [ "$#" -gt 1 ]; then
  fail "usage: install.sh [DIR]"
elif [ "$#" -eq 1 ]; then
  source_dir=$1
  [ -f "$source_dir/$name" ] || fail "$source_dir/$name was not found. Run: bun run build:all"
else
  source_dir=
  for candidate in ./dist "$here/../dist"; do
    if [ -f "$candidate/$name" ]; then
      source_dir=$candidate
      break
    fi
  done
  [ -n "$source_dir" ] || fail "$name was not found in ./dist. Run: bun run build:all, or pass the folder that holds the binaries."
fi

bin_dir=${HOME:?HOME is not set}/.local/bin
mkdir -p "$bin_dir"
cp "$source_dir/$name" "$bin_dir/teams-cli.new"
chmod 755 "$bin_dir/teams-cli.new"
# macOS marks a copied file from another source. Remove the mark so the binary starts. A missing mark is not an error.
if [ "$(uname -s)" = Darwin ]; then
  xattr -d com.apple.quarantine "$bin_dir/teams-cli.new" 2>/dev/null || true
fi
mv -f "$bin_dir/teams-cli.new" "$bin_dir/teams-cli"
echo "Installed $bin_dir/teams-cli ($platform)."

case ":${PATH:-}:" in
  *":$bin_dir:"*) ;;
  *)
    echo "$bin_dir is not on your PATH. Add this line to your shell profile (~/.zshrc or ~/.bashrc), then open a new terminal:"
    echo "  export PATH=\"\$HOME/.local/bin:\$PATH\""
    ;;
esac

echo "Next: teams-cli auth login"
