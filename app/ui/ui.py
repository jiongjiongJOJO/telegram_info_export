import flet as ft
from app.ui.main_page import MainPage
from app.ui.settings_page import SettingsPage
from app.ui.about_page import AboutPage


class FletUI(ft.Column):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True

        # 初始化页面
        self.main_page = MainPage(self.page)
        self.settings_page = SettingsPage(self.page)
        self.about_page = AboutPage(self.page)

        # 导航栏
        main_nav_rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=70,
            min_extended_width=150,
            leading=ft.Image(src="assets/logo.png", width=50, height=50),
            group_alignment=0,
            destinations=[
                ft.NavigationRailDestination(
                    icon=ft.icons.HOME_OUTLINED,
                    selected_icon=ft.icons.HOME,
                    label="主界面"
                ),
                ft.NavigationRailDestination(
                    icon=ft.icons.SETTINGS_OUTLINED,
                    selected_icon=ft.icons.SETTINGS,
                    label="设置"
                ),
                ft.NavigationRailDestination(
                    icon=ft.icons.INFO_OUTLINED,
                    selected_icon=ft.icons.INFO,
                    label="关于"
                )
            ],
            on_change=self.nav_change,
            expand=True,
        )

        # 在导航栏leading部分添加主题切换按钮
        self.theme_icon = ft.IconButton(
            icon=ft.icons.DARK_MODE if page.theme_mode == ft.ThemeMode.LIGHT else ft.icons.LIGHT_MODE,
            on_click=self.toggle_theme,
            tooltip="切换主题",
            alignment=ft.alignment.center,
        )

        self.nav_rail = ft.Column(
            [
                main_nav_rail,
                ft.VerticalDivider(width=1),
                self.theme_icon
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

        # 主内容区域
        self.content_area = ft.Container(
            content=self.main_page,
            expand=True,
            padding=ft.padding.all(20)
        )

        self.row = ft.Row(
            [self.nav_rail, ft.VerticalDivider(width=1), self.content_area],
            expand=True,
        )
        self.controls = [self.row]

    def nav_change(self, e):
        index = e.control.selected_index
        self.content_area.content = [
            self.main_page,
            self.settings_page,
            self.about_page
        ][index]
        self.page.update()

    def toggle_theme(self, e):
        if self.page.theme_mode == ft.ThemeMode.LIGHT:
            self.page.theme_mode = ft.ThemeMode.DARK
            self.theme_icon.icon = ft.icons.LIGHT_MODE
            self.main_page.qr_img.src = "assets/qr-error-light.png"
        else:
            self.page.theme_mode = ft.ThemeMode.LIGHT
            self.theme_icon.icon = ft.icons.DARK_MODE
            self.main_page.qr_img.src = "assets/qr-error-dark.png"
        self.page.update()
