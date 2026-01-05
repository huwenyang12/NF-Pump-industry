import imaplib
import email
from email.header import decode_header
import pathlib
from datetime import datetime

# 邮箱配置列表（可扩展）
MAIL_ACCOUNTS = {
    "1": {
        "name": "售后 robot1",
        "email": "db-shfw-zdrd1@nanfang-pump.com",
        "password": "3ALbQNXBex4Qwhv5",
    },
    "2": {
        "name": "售后 robot2",
        "email": "db-shfw-zdrd2@nanfang-pump.com",
        "password": "t8rY7myVKS1j3@3k",
    },
    "3": {
        "name": "售后 robot3",
        "email": "db-shfw-thd@nanfang-pump.com",
        "password": "cQ5z#Fg97vV6jsae",
    },
    "4": {
        "name": "南方流体 robot2",
        "email": "nblt-xsdd-zdrd@nanfang-pump.com",
        "password": "G@HcxFKTG91HUtTT",
    },
}

IMAP_SERVER = "imap.qiye.163.com"
IMAP_PORT = 993


def decode_str(s):
    if not s:
        return ""
    parts = decode_header(s)
    out = []
    for t, enc in parts:
        if isinstance(t, bytes):
            out.append(t.decode(enc or "utf-8", errors="ignore"))
        else:
            out.append(t)
    return "".join(out)


def save_attachment(part, outdir):
    filename = part.get_filename()
    if not filename:
        return None
    filename = decode_str(filename)
    outdir = pathlib.Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    path = outdir / filename
    with open(path, "wb") as f:
        f.write(part.get_payload(decode=True))
    return str(path)



def main():
    # ① 选择邮箱
    print("请选择要登录的邮箱：")
    for key, info in MAIL_ACCOUNTS.items():
        print(f"{key}. {info['name']} ({info['email']})")

    choice = input("\n输入序号选择邮箱：").strip()
    if choice not in MAIL_ACCOUNTS:
        print(" 无效选择。程序退出。")
        return

    account = MAIL_ACCOUNTS[choice]
    EMAIL = account["email"]
    PASSWORD = account["password"]
    print(f"\n【登录邮箱】：{account['name']} ({EMAIL})")

    # ② 选择查询日期
    input_date = input("请输入要查询的日期（格式：YYYY-MM-DD，留空=今天）：").strip()
    if input_date:
        try:
            dt = datetime.strptime(input_date, "%Y-%m-%d")
        except ValueError:
            print("日期格式错误，应为 YYYY-MM-DD，例如 2025-10-30")
            return
    else:
        dt = datetime.now()

    target_date = dt.strftime("%d-%b-%Y")
    print(f"\n【查询日期】：{target_date}\n")

    # ③ 登录 IMAP
    imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    imap.login(EMAIL, PASSWORD)
    imap.select("INBOX")

    # ④ 搜索当天邮件
    status, data = imap.search(None, f'(ON "{target_date}")')
    ids = data[0].split()

    if not ids:
        print(f"{target_date} 没有邮件。")
        imap.logout()
        return

    print(f"{target_date} 共 {len(ids)} 封邮件：\n")

    # ⑤ 输出结果 + 保存附件
    for num in reversed(ids):  # 倒序
        status, msg_data = imap.fetch(num, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        subject = decode_str(msg.get("Subject"))
        frm = decode_str(msg.get("From"))
        date = decode_str(msg.get("Date"))

        print("-----")
        print("ID:", num.decode() if isinstance(num, bytes) else num)
        print("From:", frm)
        print("Subject:", subject)
        print("Date:", date)

        # 保存附件
        for part in msg.walk():
            content_disposition = part.get("Content-Disposition", "")
            if part.get_content_maintype() == "multipart":
                continue
            if "attachment" in content_disposition.lower() or part.get_filename():
                saved = save_attachment(part, f"./attachments/{EMAIL}")
                print("📎 Saved attachment:", saved)

    imap.close()
    imap.logout()
    print("\n【查询结束】")


if __name__ == "__main__":
    main()
