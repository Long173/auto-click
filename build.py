import PyInstaller.__main__
import os
import shutil

# Tạo thư mục dist nếu chưa tồn tại
os.makedirs("dist", exist_ok=True)

# Tạo file exe với PyInstaller
PyInstaller.__main__.run([
    "pure_autoclick.py",
    "--onefile",
    "--windowed", 
    "--icon=images/autoclick.ico",
    "--name=Free Auto Clicker",
    "--add-data=images/autoclick.ico;images",
    "--add-data=user_guide.txt;.",
])

# Tạo file .bat để khởi động (tùy chọn)
with open(os.path.join("dist", "Run Auto Clicker.bat"), "w") as bat_file:
    bat_file.write('@echo off\n')
    bat_file.write('echo Dang khoi dong Auto Clicker...\n')
    bat_file.write('start "" "%~dp0Free Auto Clicker.exe"\n')

# Sao chép hướng dẫn sử dụng cho người Việt
shutil.copy("user_guide.txt", os.path.join("dist", "HUONG_DAN_SU_DUNG.txt"))

print("Build completed successfully!")
print("Executable file is in the 'dist' folder")
print("You can run the application using 'Free Auto Clicker.exe' or 'Run Auto Clicker.bat'") 