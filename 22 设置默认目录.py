"""设置工具箱文件选择窗口的默认目录。"""

from file_picker_common import pick_folder, set_default_directory


def main() -> int:
    folder = pick_folder("选择工具箱默认目录")
    if folder is None:
        return 0
    try:
        set_default_directory(folder)
    except OSError as exc:
        print(f"默认目录保存失败：{exc}")
        return 1
    print(f"工具箱默认目录已设为：{folder}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
