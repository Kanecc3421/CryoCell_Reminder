# ❄️ CryoCell Reminder

一个优雅实用的**细胞冻存提醒系统**，适用于科研人员和实验室高效管理细胞入库记录与定期提醒。支持图形化操作、批量处理、邮件发送与多冻存盒管理，具备打包为 `.exe` 独立运行的能力。

---

## 🧪 项目特色

* ✅ **图形化操作**：清晰直观的 Tkinter GUI，零代码上手
* ✅ **自定义提醒**：设置每个细胞的提醒天数，自动邮件通知
* ✅ **冻存盒管理**：支持多个冻存盒，格子位置直观可视化
* ✅ **多选操作**：支持批量添加/删除/取出细胞
* ✅ **右键菜单**：表格右键操作更方便
* ✅ **邮件功能**：配置邮箱后每天定时提醒（SMTP 支持）
* ✅ **窗口美观居中**：UI 细节打磨良好
* ✅ **支持打包为 .exe**：非开发者也能一键运行

---

## 🖥️ 软件界面

![界面截图](screenshot.png) *(如需添加，请使用实际运行截图)*

---

## 📂 项目结构

```
CryoCell_Reminder/
├── CryoCell_Reminder_v5.1_fixedpath.py     # 主程序（已支持路径修复）
├── cell_freeze.db                          # 示例数据库（运行后自动生成）
├── email_config.json                       # 邮箱配置文件（运行后设置生成）
├── ice_cube_icon.ico                       # 程序图标（小冰块 ❄️）
├── README.md                               # 本文档
└── LICENSE                                 # MIT 许可证
```

---

## 🚀 如何运行

### ✅ 方法一：直接运行 Python 代码（开发者）

1. 安装 Python 3.x
2. 安装依赖（仅 `tkinter`, `schedule`，一般已内置）
3. 双击或运行 `CryoCell_Reminder_v5.1_fixedpath.py`

### ✅ 方法二：使用已打包 `.exe` 版本（推荐）

1. 下载 `CryoCell_Reminder.exe`
2. 与以下文件放在同一目录：

   * `cell_freeze.db`
   * `email_config.json`
3. 双击 `.exe` 启动应用

---

## ✉️ 邮件配置说明

1. 第一次运行程序会弹出配置窗口：填写发件人邮箱、收件人、密码
2. SMTP 服务器默认为 `smtp.zju.edu.cn:994`（可根据需求修改）
3. 支持测试发送功能

⚠️ 密码仅保存在本地 `email_config.json`，建议使用邮件服务提供的“授权码”。

---

## ⚙️ 打包为 .exe（开发者）

确保安装 PyInstaller：

```bash
pip install pyinstaller
```

打包命令：

```bash
pyinstaller --onefile --windowed --icon=ice_cube_icon.ico CryoCell_Reminder_v5.1_fixedpath.py
```

输出文件：`dist/CryoCell_Reminder_v5.1_fixedpath.exe`

---

## 🧊 图标设计

图标采用卡通立体小冰块风格，富有科技感与实验室冷冻场景联想。

---

## ❓ 常见问题

| 问题              | 解答                                                                                    |
| --------------- | ------------------------------------------------------------------------------------- |
| `.exe` 无法读取数据库？ | 确保 `cell_freeze.db` 和 `email_config.json` 与 `.exe` 同目录，程序使用 `sys.executable` 方式定位资源路径 |
| 邮件无法发送？         | 请确认邮箱密码正确、SMTP 端口未被屏蔽，推荐使用授权码                                                         |
| 数据库会被清空吗？       | 不会。数据保存在本地 `cell_freeze.db` 中                                                         |

---

## 📜 License

MIT License. 本项目完全开源，欢迎用于学习、科研和个人使用。

---

## 🙌 致谢

本项目由 ChatGPT + 科研需求共同打磨，如你觉得有帮助欢迎 🌟 Star 本项目并分享给其他实验室同仁！
