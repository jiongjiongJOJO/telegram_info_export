import flet as ft
from flet import Page
import darkdetect
from app.ui.ui import FletUI


def main(page: Page):
    page.window.visible = False
    # 根据系统主题初始化
    initial_theme = ft.ThemeMode.DARK
    try:
        if darkdetect.theme() == "Light":
            initial_theme = ft.ThemeMode.LIGHT
    except:
        pass

    page.title = "Telegram数据导出"
    # 设置默认窗口大小
    page.window.width = 500
    page.window.height = 700

    # 设置窗口位置
    page.window.left = 100  # 左侧位置
    page.window.top = 100  # 顶部位置

    # 设置窗口最小大小
    page.window.min_width = 500  # 最小宽度
    page.window.min_height = 700  # 最小高度
    page.theme_mode = initial_theme  # 设置初始主题

    page.window.full_screen = False
    page.window.maximizable = False

    page.window.visible = True

    app = FletUI(page)
    page.add(app)


if __name__ == '__main__':
    ft.app(target=main)