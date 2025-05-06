import flet as ft
from flet import *
import requests
from app.version import VERSION


class AboutPage(ft.Column):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True
        self.scroll = "auto"

        self.content = ft.Markdown(
            """
            ### Telegram数据导出
            - 仓库地址: [https://github.com/jiongjiongJOJO/telegram_info_export](https://github.com/jiongjiongJOJO/telegram_info_export)
            - 最新版本: loading...  
            - 当前版本: v1.1.2

            ### 特别感谢
            - [PyQt5](https://pypi.org/project/PyQt5/)
            - [Telethon](https://pypi.org/project/Telethon/)
            - [Gemini](https://gemini.google.com/)
            """,
            extension_set="gitHubWeb",
            selectable=True
        )

        self.controls = [self.content]
        self.load_version()

    async def load_version(self):
        try:
            response = requests.get(
                "https://raw.githubusercontent.com/jiongjiongJOJO/telegram_info_export/master/app/version.py"
            )
            latest_version = response.text.split("=")[1].strip().strip("'")
            self.content.value = self.content.value.replace("loading...", latest_version)
        except:
            self.content.value = self.content.value.replace("loading...", "获取失败")
        finally:
            self.page.update()