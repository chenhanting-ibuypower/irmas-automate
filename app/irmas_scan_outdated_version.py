import json
import glob
from packaging import version
from app.paths import internal_path
import os

# ---------------------------
# Load policy config
# ---------------------------

POLICY_PATH = internal_path("config/software_policy.json")

with open(POLICY_PATH, "r", encoding="utf-8") as f:
    policy = json.load(f)


# ---------------------------
# Name normalization (7-Zip only)
# ---------------------------

def normalize_name(name: str) -> str:
    lower = name.lower()
    if lower.startswith("7-zip"):
        return "7-Zip"
    return name


# ---------------------------
# Rule matching logic
# ---------------------------

def matches_rule(software_name: str, rule_name: str, rule: dict) -> bool:
    match_type = rule["match_type"]

    if match_type == "keyword":
        return any(pattern.lower() in software_name.lower()
                   for pattern in rule["match_patterns"])

    if match_type == "exact":
        return software_name in rule["match_patterns"]

    if match_type == "version_threshold":
        return software_name == rule_name

    return False


# ---------------------------
# Version comparison
# ---------------------------

def is_outdated(installed_v: str, required_v: str) -> bool:
    try:
        return version.parse(installed_v) < version.parse(required_v)
    except Exception:
        return False


# ---------------------------
# Scan software inventory
# ---------------------------

SOFTWARE_DIR = os.path.join("output", "irmas", "特定軟體清查分群", "*.json")

def scan_inventory():
    rows = []

    for file in glob.glob(SOFTWARE_DIR):
        with open(file, "r", encoding="utf-8") as f:
            data = json.load(f)

        for software_name, versions_dict in data.items():
            normalized_name = normalize_name(software_name)

            for rule_name, rule in policy.items():
                if not matches_rule(normalized_name, rule_name, rule):
                    continue

                required = rule["min_required_version"]

                for installed_version, user_list in versions_dict.items():

                    if not user_list:
                        continue

                    if is_outdated(installed_version, required):

                        for u in user_list:
                            rows.append({
                                "IP位址": u.get("IP位址"),
                                "Name": u.get("使用者中文姓名"),
                                "Dept": u.get("使用者部門三"),
                                "Software": software_name,
                                "Installed": installed_version,
                                "Required": required,
                                "SourceFile": os.path.basename(file)
                            })

    return rows


# ---------------------------
# Output directory
# ---------------------------

OUTPUT_DIR = os.path.join("output", "irmas")
os.makedirs(OUTPUT_DIR, exist_ok=True)

json_path = os.path.join(OUTPUT_DIR, "outdated_softwares_detail_report.json")

# ---------------------------
# Main callable function
# ---------------------------

def run_outdated_scan():
    """
    Runs outdated software scanning based on policy rules
    and generates:
      - outdated_software.json
      - outdated_software.xlsx
    """
    rows = scan_inventory()

    print(f"Total outdated records: {len(rows)}\n")

    print(f"{'IP位址':<15} | {'Name':<10} | {'Dept':<12} | "
          f"{'Software':<25} | {'Installed':<12} | {'Required'}")
    print("-" * 96)

    for r in rows:
        print(f"{r['IP位址']:<15} | "
              f"{r['Name']:<10} | "
              f"{r['Dept']:<12} | "
              f"{r['Software']:<25} | "
              f"{r['Installed']:<12} | "
              f"{r['Required']}")

    # Save JSON
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(rows, jf, ensure_ascii=False, indent=2)


    print(f"\nSaved JSON → {json_path}")
    print("Outdated version scan completed.")


# ---------------------------
# Allow standalone execution
# ---------------------------
if __name__ == "__main__":
    run_outdated_scan()
