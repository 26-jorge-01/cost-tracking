import sys
sys.stdout.reconfigure(encoding='utf-8')
code = open(sys.argv[0], 'r', encoding='utf-8').read()
# Find the marker and extract code after it
marker = '###SCANNER_CODE_START###'
idx = code.find(marker)
if idx >= 0:
    with open(sys.argv[1], 'w', encoding='utf-8') as f:
        f.write(code[idx+len(marker):].strip())
    print(f"Written {sys.argv[1]}")
