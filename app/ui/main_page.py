import threading
import flet as ft


class MainPage(ft.Column):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True
        self.spacing = 15

        self.tab_card = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(
                    text="手机号登陆",
                    icon=ft.Icons.PHONE,
                    content=None,
                ),
                ft.Tab(
                    text='二维码登陆',
                    icon=ft.Icons.QR_CODE,
                    content=None,
                )
            ],
            expand=1,
            on_change=lambda e: self.on_tab_change(),
        )

        self.id_input = ft.TextField(label="ID")
        self.api_hash_input = ft.TextField(label="API Hash")
        self.phone_input = ft.TextField(label="手机号")
        self.qr_img = ft.Image(
            src='assets/qr-error-light.png',
            width=300,
            height=300,
            fit=ft.ImageFit.CONTAIN,
            color=ft.colors.WHITE,
            color_blend_mode=ft.BlendMode.DIFFERENCE,
            error_content=ft.Text("二维码加载失败", color=ft.colors.RED),
        )

        # 导出按钮
        self.export_btn = ft.ElevatedButton(
            "导出数据",
            icon=ft.icons.DOWNLOAD,
            on_click=self.on_export_click
        )

        # 状态显示
        self.status_label = ft.Text("", color="black")

        self.init_phone_login_ui()

        # self.init_qr_login_ui()

    def on_export_click(self, e):
        # TODO: 实现导出逻辑
        self.status_label.value = "正在获取数据..."
        self.status_label.color = "blue"
        self.page.update()

    def init_phone_login_ui(self):
        self.controls = [
            # tab_card,
            ft.Container(self.tab_card, alignment=ft.alignment.top_center),

            ft.Container(self.id_input, padding=5),
            ft.Container(self.api_hash_input, padding=5),
            ft.Container(self.phone_input, padding=5),
            ft.Container(ft.Row([self.export_btn, self.status_label]), padding=5)
        ]

    def init_qr_login_ui(self):
        self.controls = [
            # tab_card,
            ft.Container(self.tab_card, alignment=ft.alignment.top_center),

            ft.Container(self.id_input, padding=5),
            ft.Container(self.api_hash_input, padding=5),
            ft.Container(self.qr_img, padding=5),
            ft.Container(ft.Row([self.export_btn, self.status_label]), padding=5)
        ]

    def on_tab_change(self):
        selected_index = self.tab_card.selected_index
        if selected_index == 0:
            self.init_phone_login_ui()
            self.status_label.value = ''
            self.page.update()
        elif selected_index == 1:
            self.init_qr_login_ui()
            self.status_label.value = ''
            self.page.update()
