import os
import json
import pandas as pd
from datetime import datetime

class IrmasReportGenerator:
    """
    讀取 IRMAS JSON 結果，產出：
      - Excel: {output_dir}/irmas_report.xlsx
      - HTML:  {output_dir}/irmas_report_searchable.html（可收合 + 可搜尋）

    需要的 JSON 檔：
      - antivirus_detail_report.json
      - antivirus_summary.json
      - banned_softwares_detail_report.json
      - banned_softwares_report.json
      - outdated_softwares_detail_report.json
    """

    def __init__(self, output_dir: str = "output/irmas"):
        self.output_dir = output_dir

    def generate_reports(self):
        # ---------- 路徑設定 ----------
        antivirus_detail_path = os.path.join(self.output_dir, "antivirus_detail_report.json")
        antivirus_summary_path = os.path.join(self.output_dir, "antivirus_summary.json")
        banned_detail_path = os.path.join(self.output_dir, "banned_softwares_detail_report.json")
        banned_summary_path = os.path.join(self.output_dir, "banned_softwares_report.json")
        outdated_software_detail_path = os.path.join(self.output_dir, "outdated_softwares_detail_report.json")

        required_paths = [
            antivirus_detail_path,
            antivirus_summary_path,
            banned_detail_path,
            banned_summary_path,
            outdated_software_detail_path,
        ]

        if not all(os.path.exists(p) for p in required_paths):
            print("⚠️ 無法產生報表：有 JSON 檔案不存在，請先完成爬蟲產出五個 JSON。")
            for p in required_paths:
                if not os.path.exists(p):
                    print(f"  - 缺少：{p}")
            return

        # ---------- Load JSON data ----------
        with open(antivirus_detail_path, "r", encoding="utf-8") as f:
            adv_detail = json.load(f)
        with open(antivirus_summary_path, "r", encoding="utf-8") as f:
            adv_summary = json.load(f)
        with open(banned_detail_path, "r", encoding="utf-8") as f:
            ban_detail = json.load(f)
        with open(banned_summary_path, "r", encoding="utf-8") as f:
            ban_summary = json.load(f)
        with open(outdated_software_detail_path, "r", encoding="utf-8") as f:
            out_detail = json.load(f)

        # ---------- Build DataFrames ----------

        # Antivirus summary
        adv_summary_rows = [
            {
                "連線伺服器IP": row.get("value"),
                "筆數": row.get("count"),
                "明細連結": row.get("detail_link", ""),
            }
            for row in adv_summary
        ]
        df_adv_summary = pd.DataFrame(adv_summary_rows)

        # Antivirus detail (flatten)
        adv_detail_rows = []
        for server_ip, info in adv_detail.items():
            count = info.get("count")
            detail = info.get("detail") or {}
            for item in detail.get("items", []):
                r = {"連線伺服器IP": server_ip, "統計筆數": count}
                r.update(item)
                adv_detail_rows.append(r)
        df_adv_detail = pd.DataFrame(adv_detail_rows)

        # Banned softwares summary
        ban_summary_rows = [
            {
                "分類": row.get("label"),
                "軟體名稱": row.get("value"),
                "筆數": row.get("count"),
                "明細連結": row.get("detail_link", ""),
            }
            for row in ban_summary
        ]
        df_ban_summary = pd.DataFrame(ban_summary_rows)

        # Banned softwares detail
        ban_detail_rows = []
        for entry in ban_detail:
            sw_name = entry.get("value")
            for item in entry.get("items", []):
                r = {"軟體名稱": sw_name}
                r.update(item)
                ban_detail_rows.append(r)
        df_ban_detail = pd.DataFrame(ban_detail_rows)

        # Outdated softwares detail
        # 直接用原始欄位（IP位址, Name, Software, Installed, Required, Dept...）
        df_out_detail = pd.DataFrame(out_detail)

        # ---------- Export Excel ----------
        excel_path = os.path.join(self.output_dir, "irmas_report.xlsx")
        with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
            df_adv_summary.to_excel(writer, index=False, sheet_name="Antivirus_Summary")
            df_adv_detail.to_excel(writer, index=False, sheet_name="Antivirus_Detail")
            df_ban_summary.to_excel(writer, index=False, sheet_name="BannedSW_Summary")
            df_ban_detail.to_excel(writer, index=False, sheet_name="BannedSW_Detail")
            df_out_detail.to_excel(writer, index=False, sheet_name="OutdatedSW_Detail")

        print(f"✅ Excel 報表已產出：{excel_path}")

        # ---------- Export HTML (collapsible + searchable) ----------
        def make_table(df: pd.DataFrame, table_id: str) -> str:
            return df.to_html(index=False, border=0, table_id=table_id)

        adv_summary_html = make_table(df_adv_summary, "table_adv_summary")
        adv_detail_html = make_table(df_adv_detail, "table_adv_detail")
        ban_summary_html = make_table(df_ban_summary, "table_ban_summary")
        ban_detail_html = make_table(df_ban_detail, "table_ban_detail")
        out_detail_html = make_table(df_out_detail, "table_out_detail")

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        search_js = """
<script>
function filterTable(inputId, tableId) {
    var input = document.getElementById(inputId);
    var filter = input.value.toLowerCase();
    var table = document.getElementById(tableId);
    if (!table) return;
    var trs = table.getElementsByTagName("tr");

    for (var i = 1; i < trs.length; i++) {
        var rowText = trs[i].innerText.toLowerCase();
        trs[i].style.display = rowText.includes(filter) ? "" : "none";
    }
}
</script>
"""

        html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<title>IRMAS 掃描報告（搜尋 + 收合）</title>
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", Arial, sans-serif;
    margin: 20px;
}}
details {{
    margin-bottom: 20px;
    padding: 12px;
    background: #fafafa;
    border-radius: 6px;
    border: 1px solid #ddd;
}}
summary {{
    font-size: 16px;
    font-weight: bold;
    cursor: pointer;
}}
input.search-box {{
    padding: 6px 10px;
    margin: 10px 0;
    width: 40%;
    border: 1px solid #aaa;
    border-radius: 5px;
}}
table {{
    border-collapse: collapse;
    width: 100%;
    margin-top: 10px;
    font-size: 13px;
}}
th, td {{
    border: 1px solid #ccc;
    padding: 6px;
}}
th {{
    background: #eee;
}}
</style>
{search_js}
</head>
<body>

<h1>IRMAS 掃描報告</h1>
<p>產出時間：{ts}</p>

<details open>
  <summary>📌 防毒伺服器 IP 統計</summary>
  <input class="search-box" id="search_adv_summary" placeholder="搜尋防毒伺服器統計..." onkeyup="filterTable('search_adv_summary','table_adv_summary')">
  {adv_summary_html}
</details>

<details>
  <summary>🖥️ 防毒伺服器 IP → 主機明細</summary>
  <input class="search-box" id="search_adv_detail" placeholder="搜尋防毒明細..." onkeyup="filterTable('search_adv_detail','table_adv_detail')">
  {adv_detail_html}
</details>

<details open>
  <summary>📌 列管軟體統計</summary>
  <input class="search-box" id="search_ban_summary" placeholder="搜尋軟體統計..." onkeyup="filterTable('search_ban_summary','table_ban_summary')">
  {ban_summary_html}
</details>

<details>
  <summary>🖥️ 列管軟體 → 主機明細</summary>
  <input class="search-box" id="search_ban_detail" placeholder="搜尋軟體安裝明細..." onkeyup="filterTable('search_ban_detail','table_ban_detail')">
  {ban_detail_html}
</details>

<details>
  <summary>🧩 軟體版本過舊 → 主機明細</summary>
  <input class="search-box" id="search_out_detail" placeholder="搜尋需更新軟體明細..." onkeyup="filterTable('search_out_detail','table_out_detail')">
  {out_detail_html}
</details>

</body>
</html>
"""

        html_path = os.path.join(self.output_dir, "irmas_report_searchable.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)

        print(f"✅ HTML 報表已產出：{html_path}")


# Run the report generator
def generate_irmas_reports():
    generator = IrmasReportGenerator(output_dir="output/irmas")
    generator.generate_reports()
    return generator

# Example usage:
if __name__ == "__main__":
    generate_irmas_reports()