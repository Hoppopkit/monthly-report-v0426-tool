#!/usr/bin/env bash
# 建立 GitHub 倉庫、推送、啟用 GitHub Pages（做法二：獨立 repo）
set -euo pipefail

REPO_NAME="monthly-report-v0426-tool"
OWNER="Hoppopkit"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

GH="${GH:-gh}"
if ! command -v "$GH" >/dev/null 2>&1; then
  if [ -x /opt/homebrew/bin/gh ]; then
    GH=/opt/homebrew/bin/gh
  else
    echo "請先安裝 GitHub CLI: brew install gh"
    exit 1
  fi
fi

if ! "$GH" auth status >/dev/null 2>&1; then
  echo "尚未登入 GitHub。請執行："
  echo "  $GH auth login"
  echo "完成後再執行本腳本。"
  exit 1
fi

if [ ! -d .git ]; then
  git init -b main
  git add -A
  git commit -m "Initial commit: mobile MR Form V0426 tool"
fi

if git remote get-url origin >/dev/null 2>&1; then
  echo "已有 origin remote，略過 repo create，直接 push…"
  git push -u origin main
else
  "$GH" repo create "$REPO_NAME" \
    --public \
    --source=. \
    --remote=origin \
    --push \
    --description "Mobile web tool to patch MR Form V0426 (add/remove end-of-month duties)"
fi

echo "啟用 GitHub Pages（main / root）…"
PAGES_JSON='{"source":{"branch":"main","path":"/"}}'
if "$GH" api "repos/${OWNER}/${REPO_NAME}/pages" >/dev/null 2>&1; then
  echo "Pages 已存在，更新設定…"
  "$GH" api --method PUT "repos/${OWNER}/${REPO_NAME}/pages" --input - <<<"$PAGES_JSON"
else
  echo "建立 Pages 站點…"
  "$GH" api --method POST "repos/${OWNER}/${REPO_NAME}/pages" --input - <<<"$PAGES_JSON"
fi

echo ""
echo "完成！約 1–3 分鐘後可開啟："
echo "  https://${OWNER}.github.io/${REPO_NAME}/"
echo ""
echo "倉庫：https://github.com/${OWNER}/${REPO_NAME}"
