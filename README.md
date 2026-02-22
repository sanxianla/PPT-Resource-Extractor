# 📦 PPT 资源批量提取器 (PPTX Resource Extractor)

一款专为 Windows 打造的轻量、免安装的桌面小工具。可以一键**批量提取** PowerPoint 文稿（`.pptx`）中的所有无损媒体资源，并**自动分类**为图片、音频和视频文件夹。

![UI Preview](https://via.placeholder.com/600x400?text=上传到GitHub后，把这里换成你的软件截图)

## ✨ 核心特性

* 🚀 **批量处理**：支持一次性框选多个 `.pptx` 文件，程序将自动排队提取，彻底解放双手。
* 🗂️ **智能分类**：按后缀名自动识别，将资源精准放入 `图片`、`视频`、`音频` 专属文件夹，告别杂乱无章。
* ⚡ **独立运行**：单文件绿色版，**自带运行环境，无需安装 Python，无需安装 Microsoft Office**，放进 U 盘随插随用。
* 🖥️ **原生高清 UI**：调用 Windows 现代原生主题，内置高 DPI 唤醒机制，完美适配 2K/4K 屏幕，字体清晰无锯齿。
* 🛡️ **安全无损**：直接从 PPTX 底层数据流中剥离原文件，百分百保留原图/原视频的最高画质。

## 📥 下载与使用

1. 前往右侧的 **[Releases](../../releases)** 页面，下载最新的 `ppt_extractor.exe`。
2. 双击运行程序。
3. 点击 **“1. 选择多个 PPTX 文件”**（可按住 `Ctrl` 或 `Shift` 键多选）。
   > **注意**：仅支持 `.pptx` 格式。如果是老版 `.ppt` 文件，请先在 Office 中“另存为” `.pptx`。
4. 勾选你需要的资源类型（图片/视频/音频）。
5. 点击 **“2. 开始批量提取”**。提取的资源会自动保存在原 PPT 所在的目录下。

## ⚠️ 常见问题

* **提取时程序卡住或无响应？**
  请检查文件是否为 OneDrive 云端在线文件（文件图标带有 ☁️）。如果是，请先右键选择“始终保留在此设备上”下载到本地，或将文件复制到桌面上再进行提取。
* **提示权限错误或文件被占用？**
  提取前，请确保您没有在 PowerPoint 软件中打开这些目标文件，请先关闭 PPT 窗口再提取。

## 🛠️ 开发者指南 (从源码编译)

**1. 克隆项目**
```bash
git clone [https://github.com/sanxianla/PPT-Resource-Extractor.git](https://github.com/sanxianla/PPT-Resource-Extractor.git)
cd PPT-Resource-Extractor
2. 打包为 EXE
确保已安装 pyinstaller，然后在代码所在目录运行最基础的打包命令即可：
pip install pyinstaller
pyinstaller --noconsole --onefile ppt_extractor.py
(注：如果您有自定义图标 app_icon.ico，可以使用命令 pyinstaller --noconsole --onefile --icon=app_icon.ico ppt_extractor.py)
👨‍💻 关于作者
Author: 散仙

GitHub: https://github.com/sanxianla

如果这个小工具帮到了你，欢迎给这个项目点一个 ⭐️ Star！
