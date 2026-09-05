from typing import Dict

from selenium import webdriver
from selenium.webdriver.remote.remote_connection import RemoteConnection
from browserbase import Browserbase
import os

BROWSERBASE_API_KEY = "bb_live_6Q6UlQnRWatFkdOxKUz3ABabnQM"



class BrowserbaseConnection(RemoteConnection):
    session_id: str

    def __init__(self, session_id: str, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.session_id = session_id

    def get_remote_connection_headers(
        self, parsed_url: str, keep_alive: bool = False
    ) -> Dict[str, str]:
        headers = super().get_remote_connection_headers(parsed_url, keep_alive)
        headers["x-bb-api-key"] = BROWSERBASE_API_KEY
        headers["session-id"] = self.session_id
        return headers


def run() -> None:
    session = bb.sessions.create()
    connection = BrowserbaseConnection(session.id, session.selenium_remote_url)
    driver = webdriver.Remote(
        command_executor=connection, options=webdriver.ChromeOptions()
    )

    print(f"Live debug URL: https://browserbase.com/sessions/{session.id}")

    try:
        driver.get("https://www.sfmoma.org")
        print(f"At URL: {driver.current_url} | Title: {driver.title}")
    finally:
        driver.quit()


if __name__ == "__main__":
    run()
