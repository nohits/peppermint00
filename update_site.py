import os

COUNT_FILE = "count.txt"
SOURCE_HTML = "updatedtemplate.html"  # <-- read the PowerShell-updated HTML
OUTPUT_HTML = "index.html"

# Load or initialize count
if os.path.exists(COUNT_FILE):
    with open(COUNT_FILE, "r", encoding="utf-8") as f:
        count = int(f.read().strip())
else:
    count = 0

count += 1

# Save updated count
with open(COUNT_FILE, "w", encoding="utf-8") as f:
    f.write(str(count))

# Read the HTML file updated by PowerShell
with open(SOURCE_HTML, "r", encoding="utf-8") as f:
    html = f.read()

# Replace placeholder
html = html.replace("COUNT_PLACEHOLDER", str(count))

# Write to index.html
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Updated run count to {count}")
