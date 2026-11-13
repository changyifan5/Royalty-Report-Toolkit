#!/usr/bin/env python3
"""
Packer 后台登录（Playwright）。

访问 dashboard 若未登录会跳转到 /accounts/login/，填账号密码后进入
https://reports.example.invalid/zh-hans/pa/dashboard/

首次登录成功后会保存会话到 packer_auth.json，下次可用 --reuse 免密进入。

用法:
  set PACKER_EMAIL=你的邮箱
  set PACKER_PASSWORD=你的密码
  python packer_login.py

  python packer_login.py --reuse          # 用已保存会话
  python packer_login.py --headless       # 无界面
"""

from __future__ import annotations

import argparse
import getpass
import os
import sys
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

DASHBOARD_URL = "https://reports.example.invalid/zh-hans/pa/dashboard/"
SESSION_FILE = Path(__file__).resolve().parent / "packer_auth.json"


def _prompt_credentials() -> tuple[str, str]:
    email = os.environ.get("PACKER_EMAIL", "").strip()
    if not email:
        email = input("账号 / 邮箱: ").strip()
    password = os.environ.get("PACKER_PASSWORD", "")
    if not password:
        password = getpass.getpass("密码: ")
    if not email or not password:
        print("错误：需要账号和密码（环境变量 PACKER_EMAIL / PACKER_PASSWORD 或交互输入）", file=sys.stderr)
        sys.exit(1)
    return email, password


def _fill_login(page, email: str, password: str) -> None:
    page.get_by_placeholder("Account / Email Address").fill(email)
    page.get_by_placeholder("Password").fill(password)
    page.get_by_role("button", name="Sign In").click()


def login_with_password(
    *,
    email: str,
    password: str,
    headless: bool = False,
    save_session: bool = True,
) -> None:
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        context = browser.new_context(locale="zh-CN")
        page = context.new_page()

        print(f"打开 {DASHBOARD_URL}")
        page.goto(DASHBOARD_URL, wait_until="domcontentloaded")

        if "/accounts/login" in page.url:
            print("未登录，正在填写表单…")
            _fill_login(page, email, password)
        else:
            print("已有会话，无需再次登录。")

        try:
            page.wait_for_url("**/pa/dashboard/**", timeout=60_000)
        except PlaywrightTimeout:
            if page.locator(".alert-danger, .errorlist, .invalid-feedback").count():
                msg = page.locator(".alert-danger, .errorlist, .invalid-feedback").first.inner_text()
                print(f"登录失败: {msg.strip()}", file=sys.stderr)
            else:
                print(f"登录超时，当前页面: {page.url}", file=sys.stderr)
            browser.close()
            sys.exit(1)

        print(f"登录成功: {page.url}")

        if save_session:
            context.storage_state(path=str(SESSION_FILE))
            print(f"会话已保存: {SESSION_FILE}")

        if not headless:
            print("浏览器保持打开，按 Enter 关闭…")
            input()

        browser.close()


def login_with_saved_session(*, headless: bool = False) -> None:
    if not SESSION_FILE.is_file():
        print(f"未找到会话文件 {SESSION_FILE}，请先运行: python packer_login.py", file=sys.stderr)
        sys.exit(1)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        context = browser.new_context(
            locale="zh-CN",
            storage_state=str(SESSION_FILE),
        )
        page = context.new_page()

        print(f"使用已保存会话打开 {DASHBOARD_URL}")
        page.goto(DASHBOARD_URL, wait_until="domcontentloaded")

        if "/accounts/login" in page.url:
            print("会话已过期，请重新登录（不要加 --reuse）。", file=sys.stderr)
            browser.close()
            sys.exit(1)

        print(f"已进入: {page.url}")

        if not headless:
            print("浏览器保持打开，按 Enter 关闭…")
            input()

        browser.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Packer 后台登录")
    parser.add_argument("--reuse", action="store_true", help="使用 packer_auth.json 中的已保存会话")
    parser.add_argument("--headless", action="store_true", help="无界面运行")
    parser.add_argument("--no-save", action="store_true", help="登录成功后不保存会话文件")
    args = parser.parse_args()

    if args.reuse:
        login_with_saved_session(headless=args.headless)
    else:
        email, password = _prompt_credentials()
        login_with_password(
            email=email,
            password=password,
            headless=args.headless,
            save_session=not args.no_save,
        )


if __name__ == "__main__":
    main()
