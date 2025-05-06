import flet as ft
from flet import Page
from app.ui.ui import FletUI


def main(page: Page):
    page.title = "Telegram数据导出"
    page.window_width = 580
    page.window_height = 450
    page.window_resizable = False

    app = FletUI(page)
    page.add(app)


ft.app(target=main)