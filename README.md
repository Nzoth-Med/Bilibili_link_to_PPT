# Bilibili Video to PPT Converter  
*一键将Bilibili视频转换为高质量PPT | Convert Bilibili Videos to High-Quality PPT Slides*

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
