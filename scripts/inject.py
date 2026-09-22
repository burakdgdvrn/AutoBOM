import os

html_path = r'c:\Users\burak\Desktop\PDF okuma\src\ERP Sistemi.html'
with open(html_path, 'r', encoding='utf-8') as f:
    content = f.read()

if '</head>' in content and 'custom_erp.css' not in content:
    content = content.replace('</head>', '    <link rel="stylesheet" href="custom_erp.css">\n</head>')

if '</body>' in content and 'custom_erp.js' not in content:
    content = content.replace('</body>', '    <script src="custom_erp.js"></script>\n</body>')
else:
    # If no </body> tag (maybe it ends abruptly?), just append
    if 'custom_erp.js' not in content:
        content += '\n<script src="custom_erp.js"></script>'

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(content)

print('Injected custom CSS and JS tags.')
