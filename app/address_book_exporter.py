import csv
import json
import os
import codecs


class AddressBookExporter:
    URL = "https://ntpe.cht.com.tw/ldap/eo.aspx"
    OUTPUT_DIR = "output/contacts"
    OUTPUT_JSON = "address_book.json"

    def __init__(self, page):
        """
        page: Playwright sync Page (already logged in)
        """
        self.page = page

    def run(self):
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)

        csv_path = self._download_csv()
        json_path = self._convert_csv_to_json(csv_path)

        print(f"[DONE] CSV  -> {csv_path}")
        print(f"[DONE] JSON -> {json_path}")

    # ------------------------
    # Download CSV (SYNC)
    # ------------------------
    def _download_csv(self):
        page = self.page

        # Navigate only if needed
        if not page.url.startswith(self.URL):
            page.goto(self.URL, wait_until="networkidle")

        with page.expect_download() as download_info:
            page.click("#B_DOWNLOAD")

        download = download_info.value
        csv_path = os.path.join(
            self.OUTPUT_DIR,
            download.suggested_filename
        )

        download.save_as(csv_path)
        return csv_path

    # ------------------------
    # CSV (Big5) → JSON
    # ------------------------
    def _convert_csv_to_json(self, csv_path):
        json_path = os.path.join(self.OUTPUT_DIR, self.OUTPUT_JSON)

        with codecs.open(csv_path, "r", encoding="big5", errors="ignore") as f:
            reader = csv.DictReader(f)
            data = []

            for row in reader:
                data.append({
                    "full_name": row.get("姓名"),
                    "last_name": row.get("姓氏"),
                    "first_name": row.get("名字"),
                    "office": row.get("處"),
                    "company": row.get("公司"),
                    "department": self._clean(row.get("部門")),
                    "title": row.get("職稱"),
                    "fax": self._clean(row.get("傳真號碼")),
                    "business_fax": self._clean(row.get("商務傳真")),
                    "business_phone": self._clean(row.get("商務電話")),
                    "mobile": self._clean(row.get("行動電話")),
                    "company_id": row.get("公司 ID"),
                    "extension": row.get("帳戶"),
                    "email": row.get("電子郵件地址"),
                    "display_name": row.get("電子郵件顯示名稱"),
                    "category": row.get("類別"),
                })

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return json_path
    
    def _clean(self, value):
        if not value or value.strip() in ("&nbsp;", ""):
            return None
        return value.strip()
