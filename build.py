import PyInstaller.__main__
import os
import shutil
import sys

# Tạo thư mục dist nếu chưa tồn tại
if not os.path.exists('dist'):
    os.makedirs('dist')

# Thông tin ứng dụng
app_name = "Free Auto Clicker"
main_script = "pure_autoclick.py"
icon_file = None  # Có thể thêm file icon ở đây nếu bạn có

# Cấu hình đóng gói
PyInstaller.__main__.run([
    main_script,
    '--name=%s' % app_name,
    '--onefile',  # Đóng gói thành một file duy nhất
    '--windowed',  # Không hiển thị cửa sổ console
    '--clean',  # Xóa các file tạm trước khi build
    '--add-data=%s' % ("requirements.txt;."),  # Thêm file requirements.txt
    '--add-data=%s' % ("README.md;."),  # Thêm file README.md
])

print("Đã hoàn thành việc đóng gói ứng dụng!")
print(f"File .exe được tạo tại: {os.path.join('dist', app_name + '.exe')}")

# Tạo file batch để chạy ứng dụng
with open(os.path.join('dist', 'Run Auto Clicker.bat'), 'w', encoding='utf-8') as f:
    f.write('@echo off\n')
    f.write('echo Starting Free Auto Clicker...\n')
    f.write(f'"{app_name}.exe"\n')
    f.write('echo.\n')
    f.write('if %errorlevel% neq 0 (\n')
    f.write('  echo An error occurred while running the application.\n')
    f.write('  pause\n')
    f.write(')\n')

# Copy file hướng dẫn sử dụng
if os.path.exists('user_guide.txt'):
    shutil.copy2('user_guide.txt', os.path.join('dist', 'HUONG_DAN_SU_DUNG.txt'))
    print("Đã thêm file hướng dẫn sử dụng")

print("Đã tạo file batch để chạy ứng dụng")
print("Hướng dẫn phân phối:")
print("1. Chia sẻ toàn bộ thư mục 'dist' cho người dùng")
print("2. Người dùng có thể chạy file 'Run Auto Clicker.bat' hoặc trực tiếp chạy file exe") 