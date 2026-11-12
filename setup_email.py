# -*- coding: utf-8 -*-
"""JTY-BidRadar 邮箱配置向导（交互式，在本机终端运行）。

用法::

    python setup_email.py

- 授权码输入使用 getpass，不回显、不进入日志、不进入聊天记录；
- 保存到本机私密文件 data/email_secrets.json（已列入 .gitignore，绝不入仓库）；
- 授权码获取：QQ 邮箱网页版 → 设置 → 账户 → 开启「IMAP/SMTP 服务」→ 生成授权码。
"""
import getpass
import sys

import config
import email_notifier

def main() -> int:
    print("=" * 56)
    print("JTY-BidRadar 邮件通知配置（QQ SMTP）")
    print("=" * 56)
    print("授权码获取方法：QQ邮箱网页版 → 设置 → 账户 →")
    print("               开启「IMAP/SMTP服务」→ 生成授权码")
    print("注意：是授权码，不是 QQ 登录密码。\n")

    user = ""
    while not user:
        user = input("发件QQ邮箱: ").strip()
        if not user:
            print("发件邮箱不能为空，请输入。")
    to_addr = input(f"收件邮箱 [{user}]: ").strip() or user

    code = getpass.getpass("SMTP授权码（输入不回显）: ").strip()
    if not code:
        print("未输入授权码，已取消。")
        return 1
    again = getpass.getpass("再次输入授权码确认    : ").strip()
    if again != code:
        print("两次输入不一致，已取消。")
        return 1

    cfg = {
        "smtp_host": "smtp.qq.com",
        "smtp_port": 465,
        "username": user,
        "auth_code": code,
        "to_addr": to_addr,
    }
    path = email_notifier.save_config(cfg)
    print(f"\n已保存到本机私密文件：{path}")
    print("（该文件已列入 .gitignore，不会进入代码仓库；程序日志不会打印其内容）")
    print("\n下一步验证：")
    print("  python monitor.py --test-email")
    return 0


if __name__ == "__main__":
    sys.exit(main())
