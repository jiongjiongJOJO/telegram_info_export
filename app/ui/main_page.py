import flet as ft
from flet import *


class MainPage(ft.Column):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True
        self.spacing = 15

        # 输入字段
        self.id_input = ft.TextField(label="ID", width=300)
        self.api_hash_input = ft.TextField(label="API Hash", width=300)
        self.phone_input = ft.TextField(label="手机号", width=300)

        # 导出按钮
        self.export_btn = ft.ElevatedButton(
            "导出数据",
            icon=ft.icons.DOWNLOAD,
            on_click=self.on_export_click
        )

        # 状态显示
        self.status_label = ft.Text("", color="black")

        self.controls = [
            ft.Container(self.id_input, padding=5),
            ft.Container(self.api_hash_input, padding=5),
            ft.Container(self.phone_input, padding=5),
            ft.Container(self.export_btn, padding=15),
            ft.Container(self.status_label, padding=5)
        ]

    def on_export_click(self, e):
        # TODO: 实现导出逻辑
        self.status_label.value = "正在获取数据..."
        self.status_label.color = "blue"
        self.page.update()