import os
import email
from email.policy import default

directory = r"d:\ICBF\informe\mayo\correos"
files = [f for f in os.listdir(directory) if f.endswith(".eml")]

out_path = r"d:\ICBF\cost-tracking\scratch\email_summaries.txt"
with open(out_path, "w", encoding="utf-8") as out:
    out.write(f"Found {len(files)} eml files.\n")
    for f in sorted(files):
        filepath = os.path.join(directory, f)
        with open(filepath, "rb") as fp:
            msg = email.message_from_binary_file(fp, policy=default)
        
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                    break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
        
        body_clean = " ".join(body.split())[:1200] # Let's get 1200 characters to get more context
        
        out.write("=" * 60 + "\n")
        out.write(f"File: {f}\n")
        out.write(f"Subject: {msg['Subject']}\n")
        out.write(f"Date: {msg['Date']}\n")
        out.write(f"From: {msg['From']}\n")
        out.write(f"To: {msg['To']}\n")
        out.write(f"Body Preview: {body_clean}...\n")
print("Done writing email summaries.")
