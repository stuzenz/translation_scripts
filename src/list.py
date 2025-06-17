# Define the two lists
countries = [
    "LX (JA)", "AL (CN)", "AL (JA)", "BML (ID)", "CFI (NL)", "ESA (CZ SK)", 
    "ESA (PL)", "FJL (IN)", "LA (JA)", "LAP (SG)", "LAU (AU)", "LCN (CN)", 
    "LEU (DE)", "LEU (ES)", "LEU (FR)", "LEU (IT)", "LEU (NL)", "LEU (UK)", 
    "LHK (HK)", "LKR (KR)", "LMY (MY)", "LTH (TH)", "LTW (TW)", "LUS (US)", 
    "LVN (VN)", "MARS (TR)", "MIF (PH)"
]

tasks = [
    "LO - Company & Branch Setup (per entity) adjustment if required",
    "LO - Branding - Set up Letterhead & Logo Registries",
    "LO - Country-Specific Settings (e.g., default currency, language options, tax rules)",
    "LO - Entity-Specific User Roles & Permissions Adjustments",
    "LO - Local Chart of Accounts Integration/Mapping",
    "LO - Local Legal Entity Details Configuration",
    "LO - Import Country Users & attach groups",
    "LO - Configure Notification Group Registries",
    "LO - Configuration of Quote Document and Settings",
    "LO - Ensure Charge Codes have been created",
    "LO - Create company tariff Tariffs",
    "LO - Upload costs",
    "LO - Upload client rates",
    "LO - Test rates for correctness",
    "LO - Country specific CW config completed"
]

# Generate the markdown table
markdown_table = "| Country | Task |\n"
markdown_table += "|---------|------|\n"

for country in countries:
    for task in tasks:
        markdown_table += f"| {country} | {task} |\n"

# Save to a file or print
with open("country_tasks_table.md", "w") as f:
    f.write(markdown_table)

print("Markdown table generated and saved to 'country_tasks_table.md'")


