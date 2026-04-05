# EPUB 3 to EPUB 2 Converter

将 **EPUB 3** 电子书批量无损降级为 **EPUB 2** 兼容格式，**不触动原始排版**。

---

## 为什么需要它？

Calibre、Pandoc 等工具在格式转换时会重新解析并生成 HTML/CSS，复杂排版往往遭到破坏。

**本工具的策略是"最小化改动"**：不碰书籍内容和样式表，仅重构版本描述（OPF）和导航索引（NCX）。

## 核心功能

- **无损转换** — 仅修改 OPF 版本标头和导航文件，HTML/CSS 原封不动
- **目录重构** — 递归解析 EPUB 3 `nav` 节点，自动生成 EPUB 2 `toc.ncx`
- **兼容性修正** — 清理 EPUB 3 特有属性（如 `properties="nav"`），降级 OPF 版本号
- **标准封装** — `mimetype` 位于 ZIP 首位且不压缩，通过 EPUB 格式校验

## 界面特性

- **Fluent 2 设计语言** — Windows 11 Mica 磨砂背景 + 现代化卡片布局
- **拖放支持** — 直接拖入文件夹到路径栏（需安装 `tkinterdnd2`，可选）
- **实时进度** — 标题栏与进度条同步显示转换进度
- **日志时间戳** — 每条日志带 `[HH:MM:SS]` 前缀
- **子目录输出** — 可选一键输出到源目录下的 `epub2_output/` 子文件夹
- **键盘快捷键** — `Enter` 开始 / `Esc` 取消
- **主题跟随** — 自动检测系统深浅色切换并刷新 Mica 背景
- **路径记忆** — 自动保存上次使用的路径配置

---

## 快速开始

### 下载可执行文件

 [Releases](https://github.com/RRRRUDDDD/epub3_to_2/releases) 

### 从源码运行

**环境要求**：Python 3.8+

```bash
# 安装依赖
pip install lxml customtkinter

# 可选：拖放支持
pip install tkinterdnd2

# 运行
python epub3_to_2.py
```

### 使用步骤

1. **选择源目录** — 指定 EPUB 3 文件所在的文件夹（或直接拖入）
2. **设置输出位置** — 选择导出目录，或勾选"输出到源子目录"
3. **开始转换** — 点击按钮或按 `Enter`，转换日志实时滚动显示
4. **查看结果** — 完成后弹出摘要，含成功/失败计数

---

## 技术实现

| 模块 | 说明 |
|------|------|
| `EpubConverter` | 转换引擎：基于 `lxml` XPath 解析，DFS 遍历导航树，XML 安全转义 |
| `FluentGUI` | 界面层：`customtkinter` + DWM Mica API，线程安全的异步转换 |

**安全机制**：
- NCX 文本/属性值经 `xml.sax.saxutils.escape` / `quoteattr` 转义
- UTF-8 解码容错（`errors='replace'`）
- XPath 结果空值检查，友好错误提示
- mimetype 条目清除 extra 字段，严格 `ZIP_STORED`
- 修改后的 OPF 保留原始 ZipInfo 元数据（时间戳、压缩方式）
