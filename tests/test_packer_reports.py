import unittest
from unittest.mock import Mock, patch
import importlib.util
from pathlib import Path
import tempfile
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location(
    'packer_reports_20', Path(__file__).resolve().parents[1] / '20 Packer比对汇出.py'
)
p = importlib.util.module_from_spec(spec)
spec.loader.exec_module(p)


class PackerTests(unittest.TestCase):
    def test_row_uses_report_address_instead_of_composite_id_text(self):
        page = Mock(url=p.REPORTS)
        row = page.get_by_role.return_value.filter.return_value
        row.count.return_value = 1
        row.inner_text.return_value = '1001 录音 已完成比对曲库'
        self.assertIs(p.find_report_row(page, '1001'), row)
        selector = page.locator.call_args.args[0]
        self.assertIn('a[href="/zh-hans/pa/report/1001/view/"]', selector)
        page.get_by_text.assert_not_called()
        page.get_by_role.return_value.filter.assert_called_once_with(
            has=page.locator.return_value
        )

    def test_windows_profile_uses_local_appdata(self):
        with patch.object(p.sys, 'platform', 'win32'), patch.dict(p.os.environ, {'LOCALAPPDATA': '/localappdata'}):
            self.assertEqual(p.profile_dir(), Path('/localappdata/Demo-Packer/Chrome'))

    def test_state_labels(self):
        cases = {'已完成比对曲库': 'complete', '尚未比对曲库': 'idle',
                 '比對曲庫中': 'running', '正在比对曲库': 'running',
                 '比对曲库失败': 'failed', '未知': 'unknown'}
        for label, expected in cases.items():
            self.assertEqual(p.status_kind(label), expected)

    def test_submit_failure_is_not_retried(self):
        page = Mock(url=p.REPORTS)
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.get_attribute.return_value = '/zh-hans/pa/report/1001/reupload/'
        button = page.get_by_role.return_value
        button.click.side_effect = RuntimeError('network interrupted')
        with patch.object(p, 'validate_excel'), patch.object(p, 'file_digest', return_value='abc'), patch.object(p, 'uploaded_asset', return_value='old.xlsx'):
            with self.assertRaisesRegex(RuntimeError, 'network interrupted'):
                p.replace_report(page, '1001', {'upload': '/unused.xlsx', 'sha256': 'abc'}, Mock())
        self.assertEqual(button.click.call_count, 1)

    def test_ids_and_reject_invalid(self):
        self.assertEqual(p.parse_ids('1001，1002 1001'), ['1001', '1002'])
        for value in ['', '0', '1001/../../', '1001 abc']:
            with self.assertRaises(ValueError):
                p.parse_ids(value)

    def test_reject_external_action_link(self):
        with self.assertRaises(RuntimeError):
            p.site_url('https://other.example/export/')
        self.assertEqual(p.site_url('/zh-hans/pa/report/'), p.REPORTS)

    def test_pending_then_complete(self):
        page = Mock(url=p.REPORTS)
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.get_attribute.return_value = '/zh-hans/pa/report/1001/match/result/not_found/'
        with patch.object(p, 'read_status', side_effect=[('running', ''), ('complete', '')]), patch.object(p.time, 'sleep'):
            result = p.wait_complete(page, '1001', 60, 15)
        self.assertTrue(result.endswith('/1001/match/result/not_found/'))

    def test_expired_login_stops(self):
        page = Mock(url=p.BASE + '/zh-hans/accounts/login/')
        page.goto.return_value = None
        with self.assertRaisesRegex(RuntimeError, '登录状态已过期'):
            p.read_status(page, '1001')

    def test_watch_idle_does_not_start(self):
        with patch.object(p, 'read_status', return_value=('idle', '')), patch.object(p, 'start_match') as start:
            with self.assertRaisesRegex(RuntimeError, '尚未启动'):
                p.wait_state(Mock(), '1001', 0, 30)
        start.assert_not_called()

    def test_unlimited_wait_and_progress(self):
        progress = Mock()
        with patch.object(p, 'read_status', side_effect=[('running', ''), ('running', ''), ('complete', '')]), patch.object(p.time, 'sleep') as sleep:
            self.assertEqual(p.wait_state(Mock(), '1001', 0, 30, progress=progress), 'complete')
        self.assertEqual(sleep.call_count, 2)
        self.assertEqual(progress.call_count, 2)

    def test_snapshot_preserves_selected_version(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            original = folder / '报表.xlsx'
            with ZipFile(original, 'w') as z:
                z.writestr('[Content_Types].xml', '<Types/>')
                z.writestr('xl/workbook.xml', '<workbook/>')
            target, digest = p.snapshot_file(original, folder / 'job', '1001')
            original.write_bytes(b'changed later')
            self.assertEqual(p.file_digest(target), digest)
            self.assertNotEqual(p.file_digest(original), digest)

    def test_replace_waits_before_upload_then_waits_for_new_comparison(self):
        events = []
        page = Mock(url=p.REPORTS)
        task = {'id': '1001', 'mode': 'queue-replace', 'download': False}
        def wait(*args, **kwargs):
            events.append('old_done' if kwargs.get('for_replace') else 'new_done')
        def replace(*args):
            events.append('upload')
            return 'idle'
        with patch.object(p, 'read_status', return_value=('running', '')), patch.object(p, 'wait_state', side_effect=wait), patch.object(p, 'replace_report', side_effect=replace), patch.object(p, 'start_match', side_effect=lambda *a, **kw: events.append('start')), patch.object(p, 'result_link', return_value=p.REPORTS + '1001/match/result/not_found/'), patch.object(p, 'notify'):
            result = p.process(page, '1001', Path('/unused'), 0, 30, task=task)
        self.assertEqual(events, ['old_done', 'upload', 'start', 'new_done'])
        self.assertFalse(result['downloaded'])
        page.expect_download.assert_not_called()

    def test_replacement_waits_for_completion_even_if_status_returns_to_idle(self):
        states = [('running', ''), ('idle', ''), ('complete', '')]
        with patch.object(p, 'read_status', side_effect=states), patch.object(p.time, 'sleep') as sleep:
            self.assertEqual(p.wait_state(Mock(), '1001', 0, 30, for_replace=True), 'complete')
        self.assertEqual(sleep.call_count, 2)

    def test_replacement_waits_for_new_match_entry(self):
        page = Mock()
        checkpoint = Mock()
        with patch.object(p, 'start_match', side_effect=['pending', 'started']) as start, patch.object(
            p, 'read_status', side_effect=[('idle', ''), ('idle', '')]
        ), patch.object(p.time, 'sleep') as sleep:
            p.wait_to_start_replacement(page, '1001', 'complete', 0, 30, checkpoint)
        self.assertEqual(start.call_count, 2)
        self.assertEqual(sleep.call_count, 2)

    def test_notify_only_does_not_upload_start_or_download(self):
        page = Mock(url=p.REPORTS)
        task = {'id': '1001', 'mode': 'watch', 'download': False}
        with patch.object(p, 'read_status', return_value=('complete', '')), patch.object(p, 'wait_state'), patch.object(p, 'result_link', return_value=p.REPORTS + '1001/match/result/not_found/'), patch.object(p, 'start_match') as start, patch.object(p, 'replace_report') as upload, patch.object(p, 'notify'):
            result = p.process(page, '1001', Path('/unused'), 0, 30, task=task)
        self.assertEqual(result['status'], '成功')
        start.assert_not_called()
        upload.assert_not_called()
        page.expect_download.assert_not_called()

    def test_unchanged_server_file_stops_before_old_export(self):
        page = Mock(url=p.REPORTS)
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.get_attribute.return_value = '/zh-hans/pa/report/1001/reupload/'
        task = {'upload': '/unused.xlsx', 'sha256': 'abc'}
        with patch.object(p, 'validate_excel'), patch.object(p, 'file_digest', return_value='abc'), patch.object(p, 'uploaded_asset', return_value='old.xlsx'), patch.object(p, 'read_status', return_value=('complete', '')):
            with self.assertRaisesRegex(RuntimeError, '无法确认新文件'):
                p.replace_report(page, '1001', task, Mock())
        self.assertNotIn('replaced', task)
        page.expect_download.assert_not_called()

    def test_timeout_does_not_restart_match(self):
        page = Mock(url=p.REPORTS)
        with patch.object(p.time, 'monotonic', side_effect=[0, 61]):
            with self.assertRaises(TimeoutError):
                p.wait_complete(page, '1001', 60, 15)
        page.goto.assert_not_called()


if __name__ == '__main__':
    unittest.main()
