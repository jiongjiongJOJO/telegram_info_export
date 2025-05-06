import flet as ft
from flet import *


class SettingsPage(ft.Column):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.spacing = 15
        self.expand = True

        # 代理类型选择
        self.proxy_type = ft.RadioGroup(
            content=ft.Column([
                ft.Radio(value="socks5", label="SOCKS5"),
                ft.Radio(value="socks4", label="SOCKS4"),
                ft.Radio(value="http", label="HTTP"),
                ft.Radio(value="https", label="HTTPS"),
            ])
        )

        # 输入字段
        self.proxy_ip = ft.TextField(label="代理IP")
        self.proxy_port = ft.TextField(label="代理端口")
        self.username = ft.TextField(label="用户名")
        self.password = ft.TextField(label="密码", password=True)

        # 验证按钮
        self.verify_btn = ft.ElevatedButton(
            "保存并验证",
            icon=ft.icons.VERIFIED,
            on_click=self.on_verify_click
        )

        self.status_label = ft.Text("", color="black")

        self.controls = [
            ft.Text("代理设置", size=18, weight="bold"),
            ft.Divider(),
            ft.Text("代理类型:"),
            self.proxy_type,
            self.proxy_ip,
            self.proxy_port,
            self.username,
            self.password,
            self.verify_btn,
            self.status_label
        ]

    def on_verify_click(self, e):
        # TODO: 实现验证逻辑
        self.status_label.value = "正在验证代理..."
        self.status_label.color = "blue"
        self.page.update()