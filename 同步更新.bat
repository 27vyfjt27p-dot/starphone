@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =================================
echo [1/2] 正在运行 Python 转换数据...
echo =================================
:: 运行 Python 脚本，处理 Excel 并生成 JSON
python main.py

echo.
echo =================================
echo [2/2] 正在同步本地改动 (Git)...
echo =================================
:: 切换分支
git branch -M main
:: 添加所有改动到本地仓库
git add .
:: 提交改动到本地记录，如果没有改动则跳过
git commit -m "update_%date%" || echo "No changes to commit."

echo.
echo =================================
echo ✅ 本地更新已完成！
echo 💡 注意：最后一步上传已跳过，数据未同步到云端。
echo =================================
pause