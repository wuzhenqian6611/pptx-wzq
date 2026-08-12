"""pptrender.py — 用 PowerPoint COM 把 WMF/EMF 渲染为高分辨率 PNG。

PowerPoint 原生支持 WMF/EMF 的插入与渲染（比 LibreOffice 直接转换可靠，
LibreOffice 单文件转换 WMF 会输出整页空白）。

用法（内部模块，供 cli_img 调用）：
    from pptx_wzq import pptrender
    ok = pptrender.check_available()
    n_ok = pptrender.render_wmfs([(wmf_path, png_path, scale), ...])

说明：
    - 每个 WMF 放一张空白幻灯片，AddPicture 后 Slide.Export 导出 PNG；
    - scale 为放大倍数（WMF 逻辑尺寸 × scale 像素导出，公式/线条清晰）；
    - 全部渲染结束后关闭演示文稿并退出 PowerPoint。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import sys
from pathlib import Path

_DISPATCH = None


def check_available() -> bool:
    """探测 PowerPoint COM + pywin32 是否可用（不启动实例）。"""
    try:
        import pythoncom  # noqa
        import win32com.client  # noqa
    except Exception:
        return False
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")
        return True
    except Exception:
        return False


def _size_points(wmf_path: Path):
    """WMF 逻辑尺寸(px@96dpi) → PowerPoint points（1pt = 96/72 px）。"""
    try:
        from img_filter import _vector_size
        vs = _vector_size(wmf_path)
        if vs:
            return vs[0] * 72 / 96, vs[1] * 72 / 96
    except Exception:
        pass
    return 300.0, 100.0


def _trim_white(src: str, dst: str, pad: int = 8) -> None:
    """裁剪 PNG 白边：内容集中在左上角时，裁掉四周空白。"""
    from PIL import Image, ImageChops
    im = Image.open(src).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox:
        l, t, r, b = bbox
        l = max(0, l - pad)
        t = max(0, t - pad)
        r = min(im.width, r + pad)
        b = min(im.height, b + pad)
        im.crop((l, t, r, b)).save(dst)
    else:
        im.save(dst)


def render_wmfs(jobs, visible: bool = False, quiet: bool = False) -> int:
    """批量渲染。jobs: [(wmf_path, png_path, scale), ...]。

    每个 WMF 放一张空白幻灯片（左上角），整页导出后自动裁剪白边，
    得到内容区域 PNG。返回成功数量。
    """
    if not jobs:
        return 0
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        print(f"[渲染] pywin32 缺失（pip install pywin32），跳过矢量渲染：{e}",
              file=sys.stderr)
        return 0
    pythoncom.CoInitialize()
    app = None
    pres = None
    ok_n = 0
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.Visible = visible
        except Exception:
            pass
        pres = app.Presentations.Add()
        # 整页导出尺寸（幻灯片 960x720pt @2x）
        EXPORT_W, EXPORT_H = 1920, 1440
        for i, (wmf, png, scale) in enumerate(jobs, start=1):
            try:
                wmf_abs = str(Path(wmf).resolve())     # COM 需绝对路径
                png_abs = str(Path(png).resolve())
                raw_png = Path(png).with_name(
                    Path(png).stem + "_raw.png")
                w_pt, h_pt = _size_points(Path(wmf))
                slide = pres.Slides.Add(i, 12)          # ppLayoutBlank
                slide.Shapes.AddPicture(wmf_abs, 0, -1,
                                        20, 20, max(10, w_pt),
                                        max(10, h_pt))
                slide.Export(str(raw_png.resolve()), "PNG",
                             EXPORT_W, EXPORT_H)
                _trim_white(str(raw_png.resolve()), png_abs)
                raw_png.unlink(missing_ok=True)
                ok_n += 1
                if not quiet:
                    print(f"[渲染] {Path(wmf).name} → {Path(png).name}",
                          file=sys.stderr)
            except Exception as e:
                if not quiet:
                    print(f"[渲染] {Path(wmf).name} 失败: {e}",
                          file=sys.stderr)
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return ok_n


def main() -> int:  # console
    return _main()


if __name__ == "__main__":
    # 自测：渲染单张 WMF
    w = Path(sys.argv[1])
    out = Path(sys.argv[2])
    ok = render_wmfs([(w, out, 4)], quiet=False)
    print(f"渲染结果: {'成功' if ok else '失败'} → {out}")
