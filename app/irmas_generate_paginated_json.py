import json
import math
import os
from typing import Dict, Any, List

class IrmasReportMerger:
    EXPECTED_ANTIVIRUS_IP = "10.173.105.3"

    def __init__(self, base_dir: str = "output/irmas", output_subdir: str = "ready_for_dispatch"):
        self.base_dir = base_dir
        self.output_dir = os.path.join(base_dir, output_subdir)

        self.antivirus_report: Dict[str, Any] = {}
        self.banned_report: List[Any] = []
        self.outdated_report: List[Any] = []

        self.people: Dict[str, Dict[str, Any]] = {}
        self.address_book: Dict[str, Dict[str, Any]] = {}
        self.missing_contacts: List[str] = []

    # ---------------------------
    # Helpers
    # ---------------------------
    def _ensure_person(self, name: str):
        if name not in self.people:
            self.people[name] = {
                "antivirus": [],
                "bannedSoftwares": [],
                "outdatedSoftwares": [],
                "message": ""
            }

    def _normalize_name(self, name: str) -> str:
        return name.replace("　", "").strip() if name else ""
    
    # ---------------------------
    # Load files
    # ---------------------------
    def load_reports(
        self,
        antivirus_filename="antivirus_detail_report.json",
        banned_filename="banned_softwares_detail_report.json",
        outdated_filename="outdated_softwares_detail_report.json"
    ):
        with open(os.path.join(self.base_dir, antivirus_filename), encoding="utf-8") as f:
            self.antivirus_report = json.load(f)

        with open(os.path.join(self.base_dir, banned_filename), encoding="utf-8") as f:
            self.banned_report = json.load(f)

        with open(os.path.join(self.base_dir, outdated_filename), encoding="utf-8") as f:
            self.outdated_report = json.load(f)

        return self

    # ---------------------------
    # Process antivirus
    # ---------------------------
    def _process_antivirus(self):
        for reported_ip, group in self.antivirus_report.items():
            if not isinstance(group, dict) or "detail" not in group or not group["detail"]:
                continue

            for item in group["detail"].get("items", []):
                name = item.get("使用者")
                if not name:
                    continue

                self._ensure_person(name)

                entry = dict(item)
                entry["reportedIP"] = reported_ip
                entry["expectedIP"] = self.EXPECTED_ANTIVIRUS_IP
                entry["status"] = "correct" if reported_ip == self.EXPECTED_ANTIVIRUS_IP else "wrong"

                self.people[name]["antivirus"].append(entry)

    # ---------------------------
    # Process banned software
    # ---------------------------
    def _process_banned(self):
        for entry in self.banned_report:
            software = entry.get("value")
            for item in entry.get("items", []):
                name = item.get("使用者")
                if not name:
                    continue

                self._ensure_person(name)

                banned_entry = dict(item)
                banned_entry["softwareName"] = software

                self.people[name]["bannedSoftwares"].append(banned_entry)

    # ---------------------------
    # Process outdated software
    # ---------------------------
    def _process_outdated(self):
        for item in self.outdated_report:
            name = item.get("Name")
            if not name:
                continue

            self._ensure_person(name)

            outdated_entry = dict(item)
            self.people[name]["outdatedSoftwares"].append(outdated_entry)

    # ---------------------------
    # Create natural language message
    # ---------------------------
    def _build_messages(self):
        for name, data in self.people.items():

            html_parts = []

            # -----------------------------
            # Antivirus HTML message
            # -----------------------------
            av = data["antivirus"]
            if av:
                wrong = [x for x in av if x["status"] == "wrong"]

                if wrong:
                    html_parts.append("<p>你的設備有錯誤的防毒伺服器報到紀錄：</p>")
                    html_parts.append("<ul>")
                    for x in wrong:
                        html_parts.append(
                            f'<li>{x.get("電腦名稱")}（{x.get("IP位址")}） 報到於 {x.get("reportedIP")}，正確應為 {self.EXPECTED_ANTIVIRUS_IP}</li>'
                        )
                    html_parts.append("</ul>")
                else:
                    html_parts.append("<p>你的所有設備皆向正確的防毒伺服器報到。</p>")

            # -----------------------------
            # Banned software HTML message
            # -----------------------------
            banned = data["bannedSoftwares"]
            if banned:
                html_parts.append("<p>偵測到你的設備含有禁止使用的軟體：</p>")
                html_parts.append("<ul>")
                for x in banned:
                    html_parts.append(
                        f'<li>{x.get("softwareName")}（IP：{x.get("IP位址")}，電腦：{x.get("電腦名稱")}）</li>'
                    )
                html_parts.append("</ul>")

            # -----------------------------
            # Outdated software HTML message
            # -----------------------------
            outdated = data["outdatedSoftwares"]
            if outdated:
                html_parts.append("<p>你的設備有下列軟體需要更新：</p>")
                html_parts.append("<ul>")
                for x in outdated:
                    html_parts.append(
                        f'<li>{x.get("Software")}（目前 {x.get("Installed")}，需更新至 {x.get("Required")}），IP：{x.get("IP位址")}</li>'
                    )
                html_parts.append("</ul>")

            # -----------------------------
            # Combine into HTML string
            # -----------------------------
            if html_parts:
                data["message"] = "".join(html_parts)
            else:
                data["message"] = "<p>未偵測到任何問題。</p>"

    # ---------------------------
    # Public: Process all datasets
    # ---------------------------
    def process(self):
        self._process_antivirus()
        self._process_banned()
        self._process_outdated()
        self._build_messages()
        self._attach_contacts()   # 👈 NEW
        return self

    # ---------------------------
    # Pagination
    # ---------------------------
    def get_page(self, page: int = 1, page_size: int = 50) -> Dict[str, Any]:
        names = sorted(self.people.keys())
        total_pages = math.ceil(len(names) / page_size)

        start = (page - 1) * page_size
        end = start + page_size

        page_data = {name: self.people[name] for name in names[start:end]}

        return {
            "people": page_data,
            "page": page,
            "pageSize": page_size,
            "totalPages": total_pages
        }

    # ---------------------------
    # Export all pages
    # ---------------------------
    def export_pages(self, page_size: int = 50):
        os.makedirs(self.output_dir, exist_ok=True)

        names = sorted(self.people.keys())
        total_pages = math.ceil(len(names) / page_size)

        for page in range(1, total_pages + 1):
            page_data = self.get_page(page, page_size)
            out_path = os.path.join(self.output_dir, f"irmas_page_{page}.json")

            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(page_data, f, ensure_ascii=False, indent=2)

        print(f"✔ Export complete! {total_pages} pages written to {self.output_dir}")
        return self
    
    def get_name_message_list(self) -> List[Dict[str, Any]]:
        return [
            {
                "name": name,
                "message": data.get("message", ""),
                "contact": data.get("contact")
            }
            for name, data in self.people.items()
        ]

    def load_address_book(self, path="address_book.json"):
        # If absolute path OR explicit relative path, use it directly
        if os.path.isabs(path) or path.startswith("."):
            address_book_path = path
        else:
            address_book_path = os.path.join(self.base_dir, path)

        with open(address_book_path, encoding="utf-8") as f:
            raw = json.load(f)

        self.address_book = {
            self._normalize_name(p.get("full_name")): p
            for p in raw
            if p.get("full_name")
        }

        return self
    
    def _attach_contacts(self):
        self.missing_contacts = []

        for name, data in self.people.items():
            normalized = self._normalize_name(name)
            contact = self.address_book.get(normalized)

            data["contact"] = contact

            if not contact:
                self.missing_contacts.append(name)

    def export_name_message_list(self, filename="irmas_messages.json"):
        out_path = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(self.get_name_message_list(), f, ensure_ascii=False, indent=2)

        return out_path
        
    def export_missing_contacts(self, filename="missing_contacts.json"):
        out_path = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(sorted(set(self.missing_contacts)), f, ensure_ascii=False, indent=2)

        return out_path

# --------------------------- Export name and message list directly ---------------------------
def export_irmas_name_message_list(
    base_dir: str = "output/irmas",
    output_subdir: str = "ready_for_dispatch",
    filename: str = "irmas_messages.json"
) -> str:
    merger = IrmasReportMerger(base_dir, output_subdir)
    merger.load_reports().process()
    return merger.export_name_message_list(filename)

# Main execution for testing
if __name__ == "__main__":
    merger = IrmasReportMerger()
    merger.load_reports().process().export_pages(page_size=50)
    merger.export_name_message_list()    