@echo off
setlocal enabledelayedexpansion
echo Python internal network dependency installation tool.
echo Author: Shuilin
echo Date: %date%
:main_menu
echo.
echo [1] Download package for offline installation
echo [2] Install package from local directory
echo.
set /p "select=Select mode (1 or 2): "

if "!select!"=="1" (
    :retry_download
    set "package_name="
    set /p "package_name=Enter package name to download: "
    if "!package_name!"=="" (
        echo Package name cannot be empty. Please try again.
        goto retry_download
    )
    
    rem 平台选择菜单
    echo.
    echo Select target platform:
    echo 1. Linux x86
    echo 2. Linux x86_64
    echo 3. Linux ARM
    echo 4. Linux ARM
    echo 5. Linux PowerPC
    echo 6. Linux IBM Z
    echo 7. Windows x86
    echo 8. Windows x86_64
    echo 9. macOS x86_64
    echo 10. macOS ARM64
    echo 11. Linux RISC-V
    echo 12. FreeBSD x86_64
    echo 13. Solaris SPARC
    echo 14. Solaris x86_64
    echo 15. AIX PowerPC
    echo.
    
    :retry_platform
    set "platform="
    set /p "platform=Enter platform number (1-15): "
    
    rem 映射平台编号到实际平 ??
    if "!platform!"=="1" set "platform_value=manylinux2014_i686"
    if "!platform!"=="2" set "platform_value=manylinux2014_x86_64"
    if "!platform!"=="3" set "platform_value=manylinux2014_armv7l"
    if "!platform!"=="4" set "platform_value=manylinux2014_aarch64"
    if "!platform!"=="5" set "platform_value=manylinux2014_ppc64le"
    if "!platform!"=="6" set "platform_value=manylinux2014_s390x"
    if "!platform!"=="7" set "platform_value=win32"
    if "!platform!"=="8" set "platform_value=win_amd64"
    if "!platform!"=="9" set "platform_value=macosx_10_9_x86_64"
    if "!platform!"=="10" set "platform_value=macosx_11_0_arm64"
    if "!platform!"=="11" set "platform_value=manylinux2014_riscv64"
    if "!platform!"=="12" set "platform_value=freebsd_12_0_x86_64"
    if "!platform!"=="13" set "platform_value=solaris_2_11_sparc"
    if "!platform!"=="14" set "platform_value=solaris_2_11_x86_64"
    if "!platform!"=="15" set "platform_value=aix_7_2_ppc64"
    
    if not defined platform_value (
        echo Invalid platform selection. Please try again.
        goto retry_platform
    )
    
    rem Python版本输入
    :retry_python_version
    set "python_ver="
    set /p "python_ver=Enter Python version (e.g., 3.8, 3.9, 3.10, 3.11, 3.12, 3.13): "
    if "!python_ver!"=="" (
        echo Python version cannot be empty. Please try again.
        goto retry_python_version
    )
    
    mkdir "!package_name!" 2>nul
    pip download "!package_name!" --platform !platform_value! --python-version=!python_ver! --only-binary=:all: --dest "./!package_name!"
    
    echo.
    echo Download completed. Files saved in: .\!package_name!\
    echo.
    
) else if "!select!"=="2" (
    :retry_install
    set "package_name="
    set /p "package_name=Enter package name to install: "
    if "!package_name!"=="" (
        echo Package name cannot be empty. Please try again.
        goto retry_install
    )
    pip install --no-index --find-links=./"!package_name!" "!package_name!"
) else (
    echo Invalid selection. Please choose 1 or 2.
    goto main_menu
)

endlocal