# Bilibili Video to PPT Converter  
*一键将Bilibili视频转换为高质量PPT | Convert Bilibili Videos to High-Quality PPT Slides*

---

更新日志 (Changelog)
v3.1.0 - 2025-08-28

✨ 新增功能：1.在解析视频元信息时，增加实时进度显示（spinner），用户可清晰看到解析耗时。 2.合集视频分P处理

🛡️ 保持兼容性：除进度显示外，所有原有功能（去重逻辑、PPT 生成、下载策略、ffmpeg 合并等）均保持不变。

⚡ 实现方式：采用后台线程 + 终端刷新方式，不影响 yt_dlp 的正常解析与输出。

---

## 📖 Description / 项目简介

**English**  
This Python script downloads a video from Bilibili, extracts key frames, removes duplicates, and generates a high-quality PowerPoint (PPT) file.  
It is designed for educational and research purposes, where you may want to convert lecture or tutorial videos into concise slides.  

**中文**  
这个 Python 脚本可以自动从 B站 下载视频，抽取关键帧，去除重复图片，并生成一个高质量的 PPT 文件。  
适合科研、学习笔记、课堂记录等场景，将视频快速整理为简洁的幻灯片。  

---

## ✨ Features / 功能特点

- 📥 **Download** videos from Bilibili  
- 🖼️ **Extract frames** from video  
- 🔍 **Remove duplicate images** for higher PPT quality  
- 📊 **Generate PPT** with one image per slide  
- ⏱️ **Track time spent** for each step (download, extract, deduplication, PPT generation)  
- 📝 **Save run information** in `脚本运行信息.txt` in the same folder as the PPT  

---

## 📦 Installation / 安装依赖

Make sure you have Python 3.8+ installed.  
确保你已经安装 Python 3.8+。

Install required packages:  
安装所需依赖：  
```bash
pip install yt-dlp opencv-python pillow imagehash python-pptx
