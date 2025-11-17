"""
简单演示脚本：生成一个 Excel，写入测试值，并通过 Gmail 邮件发送。

使用前请修改 `GMAIL_ACCOUNT`、`GMAIL_APP_PASSWORD`、`RECIPIENTS`。
Gmail 需要开启两步验证并创建“应用专用密码”才能被脚本登录。
"""

from pathlib import Path

from openpyxl import Workbook

from autoscraper_workflow import (
    EmailSender,
    EmailSpec,
    ExcelScreenshotter,
    ScreenshotSpec,
)

EXCEL_PATH = Path("demo_report.xlsx")
SCREENSHOT_PATH = Path("demo_report.png")
GMAIL_ACCOUNT = "jacobyinzd@gmail.com"
GMAIL_APP_PASSWORD = "wodttzwwmbkjyczt"  # 16 位应用专用密码
RECIPIENTS = ["xuziheng0604@163.com"]


def build_demo_excel(path: Path) -> Path:
    """创建一个两张工作表的 Excel：Sheet1!A1=100, Sheet2!A2=300。"""

    wb = Workbook()
    sheet1 = wb.active
    sheet1.title = "Sheet1"
    sheet1["A1"] = 100

    sheet2 = wb.create_sheet("Sheet2")
    sheet2["A2"] = 300

    wb.save(path)
    return path


def capture_demo_screenshot(workbook_path: Path) -> Path:
    """用 Excel COM 截图 Sheet1 的 A1:B5 区域。"""

    screenshotter = ExcelScreenshotter()
    spec = ScreenshotSpec(
        workbook_path=workbook_path,
        sheet_name="Sheet1",
        range_address="A1:B5",
        output_path=SCREENSHOT_PATH,
    )
    return screenshotter.capture(spec)


def send_via_gmail(excel_attachment: Path, screenshot_attachment: Path) -> None:
    """用 Gmail SMTP 把 Excel + 截图作为附件发送。"""

    email_spec = EmailSpec(
        smtp_host="smtp.gmail.com",
        smtp_port=587,
        username=GMAIL_ACCOUNT,
        password=GMAIL_APP_PASSWORD,
        recipients=RECIPIENTS,
        subject="Demo Excel 报告",
        body="附件包含测试 Excel 及截图：Sheet1!A1=100, Sheet2!A2=300。",
        attachments=[excel_attachment, screenshot_attachment],
    )
    EmailSender(email_spec).send()


def main() -> None:
    excel_path = build_demo_excel(EXCEL_PATH)
    screenshot_path = capture_demo_screenshot(excel_path)
    send_via_gmail(excel_path, screenshot_path)
    print(
        "Demo 完成，Excel {} 及截图 {} 已发送给 {}。".format(
            excel_path, screenshot_path, ", ".join(RECIPIENTS)
        )
    )


if __name__ == "__main__":
    main()
