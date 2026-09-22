with open('src/ERP Sistemi.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()
start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if '<div id="hot-grid"' in line:
        start_idx = i
        break
if start_idx != -1:
    div_count = 0
    for i in range(start_idx, len(lines)):
        div_count += lines[i].count('<div')
        div_count -= lines[i].count('</div')
        if div_count <= 0 and i > start_idx:
            end_idx = i
            break
print(f'Start: {start_idx}, End: {end_idx}')
