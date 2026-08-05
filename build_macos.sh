#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "このスクリプトはmacOS上で実行してください。"
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3.11以降をインストールしてください。"
  exit 1
fi

python3 -m venv .venv-mac
source .venv-mac/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-macos.txt
python -m unittest discover -s tests -v

rm -rf build dist
python -m PyInstaller \
  --noconfirm \
  --clean \
  --onedir \
  --windowed \
  --name "個人情報変換ちゃん" \
  --osx-bundle-identifier "jp.personal-data-tools.converter" \
  --exclude-module tkinter \
  --exclude-module pandas \
  --exclude-module numpy \
  main.py

codesign --verify --deep --strict --verbose=2 "dist/個人情報変換ちゃん.app"

rm -rf "dist/dmg-stage"
mkdir -p "dist/dmg-stage"
ditto "dist/個人情報変換ちゃん.app" "dist/dmg-stage/個人情報変換ちゃん.app"
hdiutil create \
  -volname "個人情報変換ちゃん" \
  -srcfolder "dist/dmg-stage" \
  -ov \
  -format UDZO \
  "dist/個人情報変換ちゃん_Mac.dmg"
rm -rf "dist/dmg-stage"

echo ""
echo "ビルド完了:"
echo "  dist/個人情報変換ちゃん.app"
echo "  dist/個人情報変換ちゃん_Mac.dmg"
echo ""
echo "この成果物はアドホック署名です。一般配布にはDeveloper ID署名と公証を推奨します。"
