#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv

echo "=== 从 GitHub 拉取最新代码 ==="
echo ""
git status -sb
echo ""

BRANCH="$(git branch --show-current 2>/dev/null)"
if [[ -z "$BRANCH" ]]; then
  BRANCH="master"
fi

HAS_CHANGES=0
if ! git diff --quiet || ! git diff --cached --quiet; then
  HAS_CHANGES=1
elif [[ -n "$(git ls-files --others --exclude-standard)" ]]; then
  HAS_CHANGES=1
fi

DID_STASH=0
if [[ "$HAS_CHANGES" -eq 1 ]]; then
  echo "检测到本地有未提交改动，直接拉取会失败。"
  echo ""
  echo "  1 = 先暂存本地改动，拉取后再恢复（推荐）"
  echo "  2 = 取消"
  echo ""
  prompt_terminal "请选择 [1/2，直接回车=1]: " CHOICE
  CHOICE="${CHOICE:-1}"
  if [[ "$CHOICE" == "2" ]]; then
    echo "已取消。"
    pause_terminal
    exit 0
  fi
  echo ""
  echo "正在暂存本地改动..."
  if ! git stash push -u -m "拉取前自动暂存"; then
    echo "暂存失败，请手动处理后再试。"
    pause_terminal
    exit 1
  fi
  DID_STASH=1
fi

echo "正在拉取 origin ${BRANCH} ..."
echo ""
if ! git pull origin "$BRANCH"; then
  echo ""
  echo "拉取失败。"
  if [[ "$DID_STASH" -eq 1 ]]; then
    echo "正在恢复刚才暂存的本地改动..."
    git stash pop || true
  fi
  pause_terminal
  exit 1
fi

if [[ "$DID_STASH" -eq 1 ]]; then
  echo ""
  echo "正在恢复本地改动..."
  if ! git stash pop; then
    echo ""
    echo "恢复时出现冲突，请手动解决后执行：git stash drop"
    pause_terminal
    exit 1
  fi
fi

echo ""
echo "拉取完成，当前状态："
git status -sb
echo ""
pause_terminal
