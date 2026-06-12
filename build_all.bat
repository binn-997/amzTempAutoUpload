@echo off
pyinstaller --onefile --add-data "temp;temp" --name "prov_kle处理器"     prov_uploadEcel_Split_kle.py
@REM pyinstaller --onefile --add-data "temp;temp" --name "esfa_hh_hd处理器"   esfa_uploadExcel_hh_hd.py
@REM pyinstaller --onefile --add-data "temp;temp" --name "esfa_kle处理器"     esfa_uploadExcel_kle.py
@REM pyinstaller --onefile --add-data "temp;temp" --name "esfa_拆分处理器"    esfa_uploadExcel_noneSplite.py
echo 全部打包完成！
pause