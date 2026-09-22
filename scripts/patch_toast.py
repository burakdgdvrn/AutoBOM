
import re

with open('frontend/custom_erp.js', 'r', encoding='utf-8') as f:
    content = f.read()

# Insert Toast definition at the top
toast_def = '''
const Toast = Swal.mixin({
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 3000,
    timerProgressBar: true,
    didOpen: (toast) => {
        toast.onmouseenter = Swal.stopTimer;
        toast.onmouseleave = Swal.resumeTimer;
    }
});
'''
if 'const Toast = Swal.mixin' not in content:
    content = content.replace('let hot;', toast_def + '\nlet hot;')

# Replace Swal.fire with Toast.fire for simple notifications (those with 'icon' but no 'showCancelButton' or 'timer')
# Wait, some have 'timer: 2000' and 'showConfirmButton: false' which means they are already acting like toast but modal. Let's just replace all Swal.fire except those containing 'Rapor Hazýrlanýyor' or 'Emin misiniz?'

blocks = content.split('Swal.fire({')
new_content = blocks[0]
for block in blocks[1:]:
    if 'Emin misiniz?' in block or 'Rapor Haz' in block:
        new_content += 'Swal.fire({' + block
    else:
        new_content += 'Toast.fire({' + block

with open('frontend/custom_erp.js', 'w', encoding='utf-8') as f:
    f.write(new_content)
print('Patched successfully')

