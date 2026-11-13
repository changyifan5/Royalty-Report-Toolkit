#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv

finish() {
  local status="$1"
  shift
  echo ""
  echo "$@"
  printf '按 Enter 关闭窗口…'
  read -r
  exit "$status"
}

echo "=== 同步到 GitHub ==="
branch="$(git branch --show-current)" || finish 1 "无法读取当前分支。"
[[ -n "$branch" ]] || finish 1 "请先切换到要同步的分支。"
if [[ -n "$(git diff --name-only --diff-filter=U)" ]]; then
  finish 1 "仓库有未解决的冲突，请先处理。"
fi
git -c core.filemode=false status -sb

if [[ -n "$(git -c core.filemode=false status --porcelain)" ]]; then
  printf '\n这次改了什么？（直接回车 =「更新脚本」）: '
  IFS= read -r message
  git -c core.filemode=false add -A || finish 1 "暂存失败，请查看上方 Git 提示。"
  if ! git diff --cached --quiet; then
    git commit -m "${message:-更新脚本}" || finish 1 "提交失败，请查看上方 Git 提示。"
  fi
fi

git -c core.filemode=false pull --rebase origin "$branch" ||
  finish 1 "接入远端更新失败，尚未推送。若发生冲突，可解决后执行 git rebase --continue，或用 git rebase --abort 回到本地提交。"
git push origin "$branch" || finish 1 "推送失败，请查看上方 Git 提示。"
finish 0 "同步完成。"
