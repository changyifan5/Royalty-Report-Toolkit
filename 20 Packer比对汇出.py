#!/usr/bin/env python3
"""Packer 收益报表：手动登录、自动比对、等待、汇出。"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
import re
import shutil
import hashlib
import time
from collections import Counter
from datetime import datetime
from urllib.parse import urljoin, urlparse
from zipfile import is_zipfile

BASE = 'https://reports.example.invalid'
REPORTS = BASE + '/zh-hans/pa/report/'
def profile_dir():
    if sys.platform == 'win32':
        root = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData/Local')))
        return root / 'Demo-Packer/Chrome'
    if sys.platform == 'darwin':
        return Path.home() / 'Library/Application Support/Demo-Packer/Chrome'
    return Path(os.environ.get('XDG_DATA_HOME', str(Path.home() / '.local/share'))) / 'Demo-Packer/Chrome'


PROFILE = profile_dir()
DEFAULT_OUTPUT = Path.home() / 'Downloads/Packer报表'


def parse_ids(value):
    parts = re.split(r'[\s,，;；、]+', value.strip())
    if not parts or any(not re.fullmatch(r'[1-9][0-9]*', p) for p in parts):
        raise ValueError('报表 ID 只能是正整数，多个 ID 用空格或逗号分隔。')
    return list(dict.fromkeys(parts))


def site_url(href):
    url = urljoin(BASE, href)
    if urlparse(url).netloc != urlparse(BASE).netloc:
        raise RuntimeError('报表操作链接指向其他网站，已停止。')
    return url


def check_login(page):
    if '/accounts/' in page.url:
        raise RuntimeError('登录状态已过期，请重新运行并登录。')


def completed(text):
    return '已完成比对曲库' in text or '已完成比對曲庫' in text


def find_report_row(page, report_id):
    # ID 与类型可能共用一个 td（例如“1001 录音”），不能精确匹配 td 文本。
    # 按该报表唯一的查看地址定位，避免误选文件名中含有相同数字的其他行。
    view_path = f'/zh-hans/pa/report/{report_id}/view/'
    view_link = page.locator(
        f'a[href="{view_path}"], a[href="{BASE}{view_path}"]'
    )
    view_link.wait_for(state='visible')
    check_login(page)
    row = page.get_by_role('row').filter(has=view_link)
    if row.count() != 1:
        raise RuntimeError('未找到唯一对应的报表。')
    return row


def notify(report_id, message):
    print(f'\a报表 {report_id}：{message}', flush=True)
    if sys.platform == 'darwin':
        # 使用参数传文字，避免文件名或提示文字成为脚本代码。
        script = 'on run argv\n display notification (item 1 of argv) with title "Packer 报表 20" sound name "Glass"\nend run'
        try:
            subprocess.run(['/usr/bin/osascript', '-e', script, f'报表 {report_id}：{message}'],
                           timeout=10, check=False, capture_output=True)
        except (OSError, subprocess.TimeoutExpired):
            pass  # 系统通知不可用时仍保留终端提示和记录。
    elif sys.platform == 'win32':
        # Windows 自带通知区域气泡提示，无需第三方通知依赖。
        script = (
            'Add-Type -AssemblyName System.Windows.Forms; '
            'Add-Type -AssemblyName System.Drawing; '
            '$n = New-Object System.Windows.Forms.NotifyIcon; '
            '$n.Icon = [System.Drawing.SystemIcons]::Information; '
            '$n.Visible = $true; '
            '$n.BalloonTipTitle = "Packer 20"; '
            '$n.BalloonTipText = $env:PACKER_NOTICE_TEXT; '
            '$n.ShowBalloonTip(10000); Start-Sleep -Seconds 12; $n.Dispose()'
        )
        try:
            import base64
            encoded = base64.b64encode(script.encode('utf-16le')).decode('ascii')
            env = os.environ.copy()
            env['PACKER_NOTICE_TEXT'] = f'报表 {report_id}：{message}'
            subprocess.Popen(
                ['powershell.exe', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
                env=env, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        except OSError:
            pass



def status_kind(text):
    if completed(text):
        return 'complete'
    if re.search(r'比[对對].{0,8}(失败|失敗)', text):
        return 'failed'
    if re.search(r'(尚未|未开始|未開始)比[对對]', text):
        return 'idle'
    if re.search(r'(正在|等待|排队|排隊).{0,8}比[对對]|比[对對].{0,12}(中|进行|進行|排队|排隊|等待)', text):
        return 'running'
    return 'unknown'


def read_status(page, report_id):
    from playwright.sync_api import Error as BrowserError
    # 只重试只读检查。上传和启动比对绝不自动重试，防止重复提交。
    for attempt in range(6):
        try:
            response = page.goto(REPORTS + report_id + '/view/', wait_until='domcontentloaded')
            if response is not None and response.status >= 500:
                raise ConnectionError(f'网站暂时不可用（{response.status}）')
            if response is not None and response.status >= 400:
                raise RuntimeError(f'报表页面不可访问（{response.status}）')
            check_login(page)
            text = page.locator('body').inner_text()
            return status_kind(text), text
        except (BrowserError, ConnectionError):
            if page.is_closed() or attempt == 5:
                raise
            print(f'  网络检查失败，30 秒后重试（{attempt + 1}/5）…', flush=True)
            time.sleep(30)


def wait_state(page, report_id, timeout, interval, *, for_replace=False, progress=None):
    started = time.monotonic()
    while not timeout or time.monotonic() - started < timeout:
        state, text = read_status(page, report_id)
        if state == 'complete':
            return state
        if state == 'failed':
            raise RuntimeError('系统显示比对失败，请在网站查看。')
        if state == 'idle' and not for_replace:
            raise RuntimeError('报表尚未启动比对；只等待模式不会替你启动。')
        if state == 'unknown':
            raise RuntimeError('无法识别报表状态，已停止，请检查网站。')
        minutes = int((time.monotonic() - started) / 60)
        message = f'等待当前比对结束，已等待 {minutes} 分钟'
        print(f'  {report_id}：{message}…', flush=True)
        if progress:
            progress(message)
        delay = interval if not timeout else min(interval, max(0, timeout - (time.monotonic() - started)))
        time.sleep(delay)
    raise TimeoutError('等待比对超时；系统任务可能仍在进行，可稍后再次运行。')


def inspect_match_progress(page, report_id):
    page.goto(REPORTS + report_id + '/view/', wait_until='domcontentloaded')
    check_login(page)
    entry = page.get_by_text(re.compile(r'^查看比[对對]曲[库庫][进進]度$'))
    if entry.count() != 1:
        raise RuntimeError('当前报表没有唯一的「查看比对曲库进度」入口。请在比对进行时试一次。')
    before = Counter(page.locator('body').inner_text().splitlines())
    existing_pages = set(page.context.pages)
    entry.click()
    deadline = time.monotonic() + 10
    visible = []
    while time.monotonic() < deadline:
        opened = [tab for tab in page.context.pages if tab not in existing_pages]
        target = opened[0] if opened else page
        if target.url.startswith(BASE + '/'):
            lines = target.locator('body').inner_text(timeout=3000).splitlines()
            added = Counter(lines) - before if target is page else Counter(lines)
            visible = list(added.elements())
            if any(re.search(r'\d', line) for line in visible):
                break
        page.wait_for_timeout(250)
    print('\n进度入口显示的内容：', flush=True)
    print('\n'.join(visible) or '未读取到新增文字，请查看浏览器中的进度面板。', flush=True)
    input('\n查看完后按 Enter 关闭专用浏览器…')


def result_link(page, report_id):
    path = f'/zh-hans/pa/report/{report_id}/match/result/not_found/'
    link = page.locator(f'a[href="{path}"], a[href="{BASE}{path}"]')
    if link.count() != 1:
        raise RuntimeError('已完成，但找不到唯一的比对结果入口。')
    return site_url(link.get_attribute('href'))


def wait_complete(page, report_id, timeout, interval):
    wait_state(page, report_id, timeout, interval)
    return result_link(page, report_id)


def file_digest(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def validate_excel(path):
    from zipfile import ZipFile
    path = Path(path)
    if not path.is_file() or path.suffix.lower() != '.xlsx' or not is_zipfile(path):
        raise ValueError(f'需要有效的 .xlsx 文件：{path}')
    with ZipFile(path) as archive:
        if not {'[Content_Types].xml', 'xl/workbook.xml'}.issubset(archive.namelist()):
            raise ValueError(f'文件不是有效的 Excel 工作簿：{path}')


def snapshot_file(source, folder, report_id):
    source = Path(source).expanduser().resolve()
    validate_excel(source)
    target = folder / '待上传' / report_id / source.name
    target.parent.mkdir(parents=True, exist_ok=True)
    before = file_digest(source)
    shutil.copy2(source, target)
    if file_digest(target) != before or file_digest(source) != before:
        target.unlink(missing_ok=True)
        raise RuntimeError('选定文件在复制期间发生变化，请保存 Excel 后重新选择。')
    return str(target), before


def uploaded_asset(page):
    links = page.locator('a[href*="/report_file/"]')
    if links.count() != 1:
        raise RuntimeError('无法确认网站当前报表文件，已停止替换。')
    return links.get_attribute('href')


def replace_report(page, report_id, task, checkpoint):
    validate_excel(task['upload'])
    if file_digest(task['upload']) != task['sha256']:
        raise RuntimeError('待上传副本已变化，已停止。请重新创建任务。')
    old_asset = uploaded_asset(page)
    page.goto(REPORTS + report_id + '/reupload/', wait_until='domcontentloaded')
    check_login(page)
    page.locator('#id_file').set_input_files(task['upload'])
    # 记下提交边界，失败时不自动再上传一次。
    task['old_asset'] = old_asset
    checkpoint('正在提交替换；若中断请先在网站确认是否已保存')
    page.get_by_role('button', name=re.compile(r'^(?:保存[并並][取替]代[报報]表|儲存並取代報表|Save\s+and\s+replace\s+(?:the\s+)?report)$', re.IGNORECASE)).click(timeout=120_000)
    page.wait_for_url(lambda url: '/reupload/' not in url, timeout=120_000)
    state, _ = read_status(page, report_id)
    if uploaded_asset(page) == old_asset:
        raise RuntimeError('无法确认新文件已取代旧文件，不会下载旧比对结果；请在网站检查。')
    task['replaced'] = True
    checkpoint('新文件已取代旧文件')
    return state


def start_match(page, report_id, checkpoint, *, wait_for_entry=False):
    page.goto(REPORTS, wait_until='domcontentloaded')
    check_login(page)
    page.locator('#id_keyword').fill(report_id)
    page.get_by_role('button', name=re.compile('查询|查詢')).click()
    row = find_report_row(page, report_id)
    state = status_kind(row.inner_text())
    if state == 'running' or (state == 'complete' and not wait_for_entry):
        return state
    if wait_for_entry and state in ('complete', 'unknown'):
        return 'pending'
    if state != 'idle':
        raise RuntimeError('报表状态不适合启动比对，请在网站检查。')
    link = row.get_by_role('link', name=re.compile(r'^比[对對]曲[库庫]$'))
    count = link.count()
    if count == 0 and wait_for_entry:
        return 'pending'
    if count != 1:
        raise RuntimeError('找不到唯一的比对曲库入口。')
    checkpoint('正在启动比对；若中断请在网站确认状态')
    page.goto(site_url(link.get_attribute('href')), wait_until='domcontentloaded')
    check_login(page)
    checkpoint('已打开比对入口，等待完成')
    return 'started'


def wait_to_start_replacement(page, report_id, state, timeout, interval, checkpoint):
    started = time.monotonic()
    while state != 'running':
        if state == 'failed':
            raise RuntimeError('系统显示新报表处理失败，请在网站查看。')
        if state == 'idle':
            if start_match(page, report_id, checkpoint, wait_for_entry=True) != 'pending':
                return
        elapsed = time.monotonic() - started
        if timeout and elapsed >= timeout:
            raise TimeoutError('新文件已上传，等待比对入口超时。')
        checkpoint(f'新文件已上传，等待比对入口（已等待 {int(elapsed)} 秒）')
        time.sleep(min(interval, timeout - elapsed) if timeout else interval)
        state, _ = read_status(page, report_id)


def process(page, report_id, output, timeout, interval, *, task=None, checkpoint=None):
    task = task if task is not None else {'id': report_id, 'mode': 'auto', 'download': True}
    checkpoint = checkpoint or (lambda message: None)
    state, text = read_status(page, report_id)
    print(f'\n报表 {report_id}，当前状态：{state}', flush=True)
    mode = task['mode']
    if mode in ('replace', 'queue-replace'):
        checkpoint('等待可以替换报表')
        if state == 'running':
            wait_state(page, report_id, timeout, interval, for_replace=True, progress=checkpoint)
        elif state not in ('complete', 'idle', 'failed'):
            raise RuntimeError('无法识别报表状态，不能替换。')
        state = replace_report(page, report_id, task, checkpoint)
        wait_to_start_replacement(page, report_id, state, timeout, interval, checkpoint)
    elif mode == 'auto' and state == 'idle':
        start_match(page, report_id, checkpoint)
    checkpoint('等待比对完成')
    wait_state(page, report_id, timeout, interval, progress=checkpoint)
    result_url = result_link(page, report_id)
    task['result_url'] = result_url
    checkpoint('比对完成')
    print(f'\a  {report_id} 比对完成：{result_url}', flush=True)
    page.goto(result_url, wait_until='domcontentloaded')
    check_login(page)
    summary = page.locator('body').inner_text()
    if not task['download']:
        screenshot = output / f'{report_id}_比对完成_{datetime.now():%Y%m%d_%H%M%S_%f}.png'
        page.screenshot(path=str(screenshot), full_page=False)
        notify(report_id, '比对完成，可以查看结果；未汇出')
        checkpoint('已提醒，未汇出')
        return {'id': report_id, 'status': '成功', 'downloaded': False,
                'result_url': result_url, 'summary': summary, 'screenshot': str(screenshot)}
    page.locator('#export-report').wait_for(state='visible')
    checkpoint('正在汇出')
    with page.expect_download(timeout=(timeout or 86400) * 1000) as event:
        page.locator('#export-report').click()
    download = event.value
    filename = Path(download.suggested_filename).name
    if not filename.lower().endswith('.xlsx'):
        raise RuntimeError(f'系统返回的文件不是 Excel：{filename}')
    target = output / f'{report_id}_{datetime.now():%Y%m%d_%H%M%S_%f}_{filename}'
    temporary = target.with_suffix('.partial')
    try:
        download.save_as(temporary)
        if not is_zipfile(temporary):
            raise RuntimeError('下载内容不是有效的 xlsx 文件。')
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)
    page.screenshot(path=str(target.with_suffix('.png')), full_page=False)
    notify(report_id, '比对完成，Excel 已下载')
    checkpoint('已下载')
    print(f'  完成：{target}', flush=True)
    return {'id': report_id, 'status': '成功', 'file': str(target), 'summary': summary}


def main(argv=None):
    parser = argparse.ArgumentParser(description='Packer 自动比对与汇出（不汇入、不结算）')
    parser.add_argument('ids', nargs='*', help='报表 ID，支持空格或逗号分隔')
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument('--mode', choices=['auto', 'replace', 'watch', 'queue-replace'], help='不指定时交互选择')
    parser.add_argument('--status-file', type=Path, help='供报表进度台读取的任务状态文件')
    parser.add_argument('--yes-download', action='store_true', help='直接下载 Excel，无需确认')
    parser.add_argument('--notify-only', action='store_true', help='完成后只提醒，不下载')
    parser.add_argument('--inspect-progress', action='store_true', help='只打开一个报表的进度面板并显示文字')
    parser.add_argument('--file', type=Path, help='替换文件（仅一个报表时可用）')
    parser.add_argument('--timeout', type=int, default=0, help='比对等待秒数，0 为持续等待（默认）')
    parser.add_argument('--interval', type=int, default=30, help='状态检查间隔秒数')
    args = parser.parse_args(argv)
    if args.timeout < 0 or args.interval <= 0:
        parser.error('等待时间不能小于 0，检查间隔必须大于 0')
    try:
        ids = parse_ids(' '.join(args.ids) if args.ids else input('报表 ID（多个用空格或逗号分隔）：'))
    except ValueError as exc:
        parser.error(str(exc))
    if args.inspect_progress and (len(ids) != 1 or args.mode or args.file or args.notify_only):
        parser.error('--inspect-progress 只接受一个报表 ID，不与其他操作模式同时使用')
    if args.yes_download and (not args.mode or args.notify_only or args.inspect_progress):
        parser.error('--yes-download 需要指定运行模式，并且不能与仅提醒或查看进度同时使用')
    if args.status_file and not args.yes_download:
        parser.error('--status-file 需要 --yes-download')
    mode = args.mode
    if mode is None and not args.inspect_progress:
        print('\n这次要做什么？\n')
        print('  1  换一份新报表，再比对')
        print('     现在选好新文件；如果旧比对还没结束，工具会先等完再替换。\n')
        print('  2  等正在进行的比对完成')
        print('     使用网站现有报表，不上传新文件；已经完成的直接查看结果。\n')
        print('  3  开始比对网站现有报表')
        print('     不换文件；未开始就启动，正在进行就继续等。\n')
        choice = input('输入 1、2 或 3：').strip()
        mode = {'1': 'replace', '2': 'watch', '3': 'auto'}.get(choice)
        if mode is None:
            parser.error('请输入 1、2 或 3')
    download = not args.notify_only
    if not args.notify_only and not args.inspect_progress and not args.yes_download:
        print('\n比对完成后，要下载 Excel 吗？')
        print('  Y  下载 Excel，并提醒我')
        print('  N  只提醒我，我自己到结果页查看错误')
        choice = input('输入 Y 或 N（直接回车默认下载）：').strip().lower()
        if choice not in ('', 'y', 'n'):
            parser.error('请输入 Y 或 N')
        download = choice != 'n'
    if args.file and (len(ids) != 1 or mode not in ('replace', 'queue-replace')):
        parser.error('--file 仅用于单个报表的替换模式')
    output = args.output.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    run_dir = output / f'任务_{datetime.now():%Y%m%d_%H%M%S_%f}'
    run_dir.mkdir(mode=0o700)
    tasks = []
    for report_id in ids:
        task = {'id': report_id, 'mode': mode, 'download': download, 'phase': '待处理'}
        if mode in ('replace', 'queue-replace'):
            if args.file:
                source = args.file
            else:
                from file_picker_common import pick_excel_file
                source = pick_excel_file(f'为报表 {report_id} 选择要替换上传的新文件（.xlsx）')
                if source is None:
                    print('已取消；尚未操作网站。')
                    return 0
            task['source'] = str(source)
            task['upload'], task['sha256'] = snapshot_file(source, run_dir, report_id)
            print(f'报表 {report_id} 将上传：{source}（已保留本次副本）')
        tasks.append(task)
    log = run_dir / '任务记录.json'
    def save():
        data = json.dumps(tasks, ensure_ascii=False, indent=2)
        for destination in (log, args.status_file):
            if destination is not None:
                temporary = destination.with_suffix('.tmp')
                temporary.write_text(data, encoding='utf-8')
                temporary.replace(destination)
    save()
    print(f'\n任务记录：{log}\n保持电脑唤醒、网络连接，勿关闭工具及专用 Chrome。', flush=True)
    from playwright.sync_api import sync_playwright
    PROFILE.mkdir(parents=True, exist_ok=True, mode=0o700)
    os.chmod(PROFILE, 0o700)
    results = []
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            str(PROFILE), channel='chrome', headless=False,
            accept_downloads=True, locale='zh-CN',
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(30_000)
            # 不自动接受网站确认框、协议或验证码。
            page.goto(REPORTS, wait_until='domcontentloaded')
            if '/accounts/' in page.url:
                print('请在新开的专用 Chrome 窗口完成登录（密码不用填在这里）。')
                if args.status_file:
                    for task in tasks:
                        task['phase'] = '请在专用 Chrome 完成登录'
                    save()
                    page.wait_for_url(
                        lambda url: url.startswith(BASE + '/zh-hans/pa/'), timeout=0,
                        wait_until='domcontentloaded',
                    )
                else:
                    input('登录成功后，回到这里按 Enter 继续…')
                page.goto(REPORTS, wait_until='domcontentloaded')
                check_login(page)
            if args.inspect_progress:
                inspect_match_progress(page, ids[0])
                return 0
            for task in tasks:
                report_id = task['id']
                def checkpoint(message):
                    task['phase'] = message
                    task['updated_at'] = datetime.now().isoformat(timespec='seconds')
                    save()
                try:
                    result = process(page, report_id, output, args.timeout, args.interval,
                                     task=task, checkpoint=checkpoint)
                except Exception as exc:
                    print(f'  {report_id} 失败：{exc}', flush=True)
                    notify(report_id, '任务中断，请查看终端错误及网站状态')
                    result = {'id': report_id, 'status': '失败', 'error': str(exc)}
                    try:
                        page.screenshot(path=str(output / f'{report_id}_失败_{datetime.now():%Y%m%d_%H%M%S}.png'))
                    except Exception:
                        pass
                results.append(result)
                task['result'] = result
                if result['status'] == '失败':
                    task['phase'] = '失败，请查看错误及网站状态后重新创建任务'
                save()
            if not download and any(r['status'] == '成功' for r in results):
                print('\a已完成的结果链接已保存在任务记录；当前页面保持打开，可查看错误类型。')
                input('查看完后按 Enter 关闭专用浏览器…')
        finally:
            context.close()
    successes = sum(r['status'] == '成功' for r in results)
    print(f'\n处理结束：成功 {successes} / {len(ids)}。文件夹：{output}\n运行记录：{log}\a', flush=True)
    return 0 if successes == len(ids) else 1


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print('\n已停止。网站已启动的任务可能仍会继续，下次运行会检查状态。')
        raise SystemExit(130)
    except Exception as exc:
        print(f'无法运行：{exc}\n请确认 Chrome 已安装，且没有另一个本工具正在运行。')
        raise SystemExit(1)
