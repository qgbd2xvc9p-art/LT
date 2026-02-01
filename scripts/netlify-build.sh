#!/usr/bin/env bash
set -euo pipefail

FLUTTER_VERSION="${FLUTTER_VERSION:-stable}"
FLUTTER_HOME="$HOME/flutter"

if [ ! -d "$FLUTTER_HOME/.git" ]; then
  git clone --depth 1 -b "$FLUTTER_VERSION" https://github.com/flutter/flutter.git "$FLUTTER_HOME"
else
  git -C "$FLUTTER_HOME" fetch --depth 1 origin "$FLUTTER_VERSION"
  git -C "$FLUTTER_HOME" checkout "$FLUTTER_VERSION"
fi

export PATH="$FLUTTER_HOME/bin:$PATH"

flutter --version
flutter config --enable-web
flutter pub get
flutter build web --pwa-strategy=none --dart-define=APP_VERSION="${COMMIT_REF:-local}"
