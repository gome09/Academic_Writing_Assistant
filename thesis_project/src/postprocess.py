# -*- coding: utf-8 -*-
"""Word COM 后处理：静默更新目录/页码等域，可选导出 PDF。

仅 Windows + 本机安装 Word 时可用；任何失败都只打印提示、返回 False，
不影响主流程（docx 已带 updateFields 标记作为兜底）。
"""
from __future__ import annotations
import os

_WD_EXPORT_PDF = 17  # WdExportFormat.wdExportFormatPDF


def refresh_word_fields(docx_path: str, export_pdf: bool = False) -> bool:
    """用本机 Word 更新 docx 全部域（目录/页码/交叉引用）并保存。"""
    try:
        import win32com.client
    except ImportError:
        print("  [提示] 未安装 pywin32，跳过域刷新"
              "（打开 Word 后按提示更新域即可）：pip install pywin32")
        return False
    path = os.path.abspath(docx_path)
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone：无人值守，禁止弹窗
        doc = word.Documents.Open(path, ConfirmConversions=False,
                                  AddToRecentFiles=False)
        doc.Fields.Update()
        for i in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(i).Update()
        doc.Save()
        print("  ✔ 已用 Word 刷新目录/页码域")
        if export_pdf:
            try:
                pdf = os.path.splitext(path)[0] + ".pdf"
                doc.ExportAsFixedFormat(pdf, _WD_EXPORT_PDF)
                print(f"  ✔ PDF : {pdf}")
            except Exception as e:  # noqa: BLE001
                print(f"  [提示] PDF 导出失败（{e}），域刷新已完成")
        return True
    except Exception as e:  # noqa: BLE001 COM 异常类型不可枚举
        print(f"  [提示] Word 域刷新失败（{e}），请打开文档后按 F9 手动更新")
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:  # noqa: BLE001
            pass
