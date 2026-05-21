import os
import email
from email.policy import default

directory = r"d:\ICBF\informe\mayo\correos"
files = [f for f in os.listdir(directory) if f.endswith(".eml")]

print(f"Found {len(files)} eml files.")
for f in sorted(files):
    filepath = os.path.join(directory, f)
    with open(filepath, "rb") as fp:
        msg = email.message_from_binary_file(fp, policy=default)
    print("=" * 60)
    print(f"File: {f}")
    print(f"Subject: {msg['Subject']}")
    print(f"Date: {msg['Date']}")
    print(f"From: {msg['From']}")
    print(f"To: {msg['To']}")
