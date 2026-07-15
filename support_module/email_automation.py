import os
import re
from datetime import datetime
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

def safe_filename(name: str) -> str:
    # 移除 Windows 不允許的檔名字元
    return re.sub(r'[\\/*?:"<>|]+', "_", name)


def download_latest_matching_attachments(
    subject_contains: str,
    save_dir: str,
    mailbox: str | None = None,   # None = default mailbox; 或填 "your.name@company.com"
    folder_path: str = "收件匣",   # 中文 Outlook 常見為「收件匣」；英文通常是 "Inbox"
    allowed_ext: tuple[str, ...] | None = None,  # 例如 (".pdf", ".xlsx")
    max_items_scan: int = 200
) -> list[str]:
    """
    找到最新一封主旨包含 subject_contains 的信，下載其附件到 save_dir。
    回傳已儲存檔案路徑列表。
    """
    os.makedirs(save_dir, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
    for account in accounts:
        print(account, type(account.DeliveryStore.DisplayName))


    # 選 mailbox（可選）
    if mailbox:
        root = outlook.Folders.Item(mailbox)
        folder = root.Folders.Item(folder_path)
    else:
        # 預設信箱的 folder
        folder = outlook.GetDefaultFolder(6)  # 6 = Inbox
        print(folder.Name)
        # 如果你想指定非 Inbox 的資料夾，可用 folder.Folders.Item("子資料夾名")

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # 最新在前

    saved = []

    scanned = 0
    for mail in items:
        print(f"Scanning email: {getattr(mail, 'Subject', 'No Subject')}")
        scanned += 1
        if scanned > max_items_scan:
            break

        # 有些 item 不是 MailItem（例如 meeting request）
        try:
            subject = str(mail.Subject)
        except Exception:
            continue

        if subject_contains.lower() not in subject.lower():
            continue

        # 找到最新符合條件的那封
        received = getattr(mail, "ReceivedTime", None)
        received_str = ""
        if received:
            # Outlook 的 ReceivedTime 是 COM datetime
            try:
                received_str = received.strftime("%Y%m%d_%H%M%S")
            except Exception:
                received_str = datetime.now().strftime("%Y%m%d_%H%M%S")

        attachments = mail.Attachments
        if attachments.Count == 0:
            return []

        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            filename = safe_filename(att.FileName)

            # 副檔名過濾（可選）
            if allowed_ext:
                ext = os.path.splitext(filename)[1].lower()
                if ext not in tuple(e.lower() for e in allowed_ext):
                    continue

            # 避免同名覆蓋：加上時間戳
            base, ext = os.path.splitext(filename)
            final_name = f"{base}__{received_str}{ext}" if received_str else filename
            final_path = os.path.join(save_dir, final_name)

            att.SaveAsFile(final_path)
            saved.append(final_path)

        return saved  # 只抓最新一封符合條件就結束

    return []


if __name__ == "__main__":
    files = download_latest_matching_attachments(
        subject_contains="stock cover",      # 例："CMA CGM" 或 "Track & Trace"
        save_dir=r"YOUR_NETWORK_DRIVE\ShipmentTracking\DSV進櫃計畫表",
        mailbox=None,                     # 若有多信箱才需要填
        folder_path="收件匣",             # 英文 Outlook 用 "Inbox"；或直接用預設 GetDefaultFolder(6) 不理它
        allowed_ext=(".xlsm",),                 # 例：(".pdf", ".xlsx")
        max_items_scan=300
    )
    print("Saved files:")
    for f in files:
        print(" -", f)