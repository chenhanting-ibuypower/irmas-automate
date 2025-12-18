import os
import pwinput
from dotenv import load_dotenv
from smartcard.System import readers
from playwright.sync_api import Page, TimeoutError

class ChtSsoLogin:
    """
    Generic CHT SSO login helper.

    Usage:
        login = ChtSsoLogin()
        login.ensure_login(page, "https://irmas.cht.com.tw")
        login.ensure_login(page, "https://masis.cht.com.tw/masis/Menu.aspx")
    """

    def __init__(self):
        load_dotenv()

        # Keep original env names so it doesn't break your setup
        self.account = os.getenv("EMS_ACCOUNT")
        # Note: EMS_PPASSWORD has a double P in your original code; keep it for compatibility
        self.password = os.getenv("EMS_PPASSWORD") or os.getenv("EMS_PASSWORD")
        self.card_password = os.getenv("EMS_CARD_PASSWORD")

    # -----------------------------
    # 🔍 Smart card detection
    # -----------------------------
    def check_card_presence(self) -> bool:
        """Returns True if a smart card is inserted and readable."""
        r = readers()
        if not r:
            print("❌ No smart card reader found.")
            return False

        reader = r[0]
        print(f"Using reader: {reader}")

        connection = reader.createConnection()
        try:
            connection.connect()
            print("✔ Smart card detected")
            return True
        except Exception:
            print("❌ No card inserted")
            return False

    # -----------------------------
    # 🔎 Detect if current page is SSO login
    # -----------------------------
    def _is_login_page(self, page: Page) -> bool:
        """
        Heuristic: if the SSO username input exists, we treat it as login page.
        Adjust selector if your IdP changes.
        """
        try:
            page.wait_for_selector('input[name="username"]', timeout=6000)
            return True
        except TimeoutError:
            return False

    # -----------------------------
    # 🟦 Smart card login flow
    # -----------------------------
    def _card_login(self, page: Page, account: str):
        print("🔐 使用識別證登入模式")

        # Step 1️⃣: username step (may be skipped by SSO)
        if self._has_username_input(page):
            print("🧑 偵測到帳號輸入頁，填入帳號")
            page.fill('input[name="username"]', account)
            page.click('input[name="login"]')

            # Switch login method if needed
            if page.locator("#try-another-way").count() > 0:
                page.click("#try-another-way")
                page.get_by_text("識別證登入").click()
        else:
            print("ℹ️ 未偵測到帳號頁，SSO 直接進入卡片驗證")

        # Step 2️⃣: card PIN
        page.wait_for_selector("input[name='card_pin']", timeout=10000)
        card_password = self.card_password or pwinput.pwinput(
            "請輸入您的卡片密碼(通常是身分證後8碼): ", mask="*"
        )

        page.fill("input[name='card_pin']", card_password)
        page.click("#verify-button")

        page.wait_for_timeout(10000)
        page.wait_for_load_state("networkidle")

    # -----------------------------
    # 🟨 Account + Password + OTP login flow
    # -----------------------------
    def _password_otp_login(self, page: Page, account: str):
        print("🔐 使用『帳號 + 密碼 + OTP』登入模式")

        password = self.password or pwinput.pwinput(
            "請輸入您的密碼: ", mask="*"
        )

        # Fill account & go
        page.fill('input[name="username"]', account)
        page.click('input[name="login"]')

        # Switch login method
        page.click("#try-another-way")
        page.get_by_text("OTP驗證").click()

        page.fill("input[name='password']", password)
        page.click("#kc-login")

        # OTP loop
        while True:
            received = input("📨 是否已收到 OTP？(Y/N): ").strip().lower()
            if received == "y":
                break
            elif received == "n":
                print("🔁 重新發送 OTP...")
                page.click("#kc-sendotp")
            else:
                print("❌ 請輸入 Y 或 N。")

        otp = input("請輸入 OTP 動態密碼: ")
        page.fill("input[name='sms_otp']", otp)
        page.click("#kc-login")

        page.wait_for_load_state("networkidle")

    # -----------------------------
    # 🔓 Main: ensure login for any URL
    # -----------------------------
    def ensure_login(self, page: Page, url: str) -> Page:
        """
        Go to the given URL. If redirected to SSO login, perform login.
        If already logged in (no login form detected), do nothing.

        Parameters:
            page: Playwright Page
            url: Target system URL (IRMAS, MASIS, SPAS, etc.)
        """
        print(f"🌐 Navigating to: {url}")
        page.goto(url)
        page.wait_for_load_state("networkidle")

        # Check if we see login page
        if not self._is_login_page(page):
            print("✔ 看起來已經登入，未偵測到 SSO 登入畫面。")
            return page

        print("🔑 偵測到 SSO 登入畫面，開始登入流程...")

        account = self.account or input("請輸入您的帳號: ")
        has_card = self.check_card_presence()

        if has_card:
            self._card_login(page, account)
        else:
            self._password_otp_login(page, account)

        print("✔ Login successful!")
        return page

    def _has_username_input(self, page: Page) -> bool:
        return page.locator('input[name="username"]').count() > 0
