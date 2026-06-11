简介
•	项目名：Batch Printing Tool for Multiple Documents（多格式文档批量打印工具）
•	简要：基于 Windows 的 GUI 工具，使用系统默认关联软件按顺序将多个文档发送到指定打印机。适用于 Word/Excel/PDF/图片/文本等常见格式。
主要特性
•	支持按规则对文件名排序并批量打印（支持注：N-M 类似命名自动排序）。
•	使用系统默认程序执行“打印”操作，保证原始文件格式与打印设置一致。
•	可从 UI 选择目录、筛选文件类型、多选文件并查看打印进度与日志。
•	可加载并选择系统打印机，保存最近使用的打印机配置（~/.document_printer/settings.json）。
•	自动检测并在非管理员权限下以管理员身份重启以便设置默认打印机（Windows 权限问题处理）。
•	对常见错误做出友好提示（例如：未安装默认打开程序时的提示）。
支持格式（非穷尽）
•	文档：.doc .docx .xls .xlsx .ppt .pptx .pdf
•	文本/日志：.txt .log
•	图片：.png .jpg .jpeg .ico .img
依赖（Windows）
•	Python 3.x（推荐 3.8+）
•	pywin32（win32print）
•	tkinter（随大部分 Windows Python 分发包含）
•	可选：pyinstaller（打包为可执行文件）
快速运行
1.	安装依赖：
•	pip install pywin32
2.	运行：
•	python zuoyedaying.py
3.	在界面中选择目录、筛选类型、加载打印机、选中文件并点击“打印选中文件”。
可打包为单文件（可选）
•	使用 PyInstaller：pyinstaller --onefile zuoyedaying.py
•	打包时注意包含 tkinter 相关资源和 pywin32 模块（参考 PyInstaller 文档）。
常见问题与排查
•	“系统未找到该格式的默认打开程序”：为该文件类型设置默认程序或安装相应查看/编辑软件。
•	打印未发送：确认打印机已在线且驱动正常；以管理员权限运行以允许设置默认打印机。
•	无法加载打印机列表：检查系统打印服务与网络打印机连接权限。
配置与日志
•	保存位置：用户主目录下 .document_printer/settings.json
•	日志：应用内“操作日志”窗口，运行后可手动清空。
许可证
•	本项目包含 LICENSE.txt（GNU Affero General Public License v3）。
贡献
•	欢迎提交 issue 与 PR。请在提交前确保描述复现步骤及运行环境（Windows 版本、Python 版本、打印机型号）。
