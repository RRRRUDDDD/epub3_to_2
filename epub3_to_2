import zipfile
import re
import threading
import logging
from pathlib import Path
from typing import List, Dict, Final
from lxml import etree
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

class EpubConverter:
    NAMESPACES: Final[Dict[str, str]] = {
        'n': 'urn:oasis:names:tc:opendocument:xmlns:container',
        'opf': 'http://www.idpf.org/2007/opf',
        'xhtml': 'http://www.w3.org/1999/xhtml',
        'epub': 'http://www.idpf.org/2007/ops',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'ncx': 'http://www.daisy.org/z3986/2005/ncx/'
    }

    def __init__(self):
        self._play_order = 1

    def _get_metadata(self, opf_xml: etree._Element, tag: str) -> str:
        results = opf_xml.xpath(f'//dc:{tag}/text()', namespaces=self.NAMESPACES)
        return str(results[0]).strip() if results else ""

    def _parse_nav_level(self, element: etree._Element, nav_rel_dir: Path) -> List[Dict]:
        items = []
        for li in element.xpath('./xhtml:li', namespaces=self.NAMESPACES):
            a_tags = li.xpath('./xhtml:a', namespaces=self.NAMESPACES)
            if not a_tags: continue
            a = a_tags[0]
            text = "".join(a.itertext()).strip()
            href = a.get('href')
            src = (nav_rel_dir / href).as_posix() if nav_rel_dir.name else href
            item = {'text': text, 'src': src, 'children': []}
            sub_ols = li.xpath('./xhtml:ol', namespaces=self.NAMESPACES)
            if sub_ols:
                item['children'] = self._parse_nav_level(sub_ols[0], nav_rel_dir)
            items.append(item)
        return items

    def _build_ncx_points(self, nav_items: List[Dict]) -> str:
        xml_parts = []
        for item in nav_items:
            p_id = f"nav_{self._play_order}"
            xml = f'<navPoint id="{p_id}" playOrder="{self._play_order}">\n'
            xml += f'  <navLabel><text>{item["text"]}</text></navLabel>\n'
            xml += f'  <content src="{item["src"]}"/>\n'
            self._play_order += 1
            if item['children']:
                xml += self._build_ncx_points(item['children'])
            xml += '</navPoint>\n'
            xml_parts.append(xml)
        return "".join(xml_parts)

    def convert(self, input_path: Path, output_path: Path):
        self._play_order = 1
        with zipfile.ZipFile(input_path, 'r') as zin:
            container_xml = etree.fromstring(zin.read('META-INF/container.xml'))
            opf_rel_path = container_xml.xpath('//n:rootfile/@full-path', namespaces=self.NAMESPACES)[0]
            opf_path = Path(opf_rel_path)
            opf_xml = etree.fromstring(zin.read(opf_path.as_posix()))
            
            title = self._get_metadata(opf_xml, 'title') or "Untitled"
            author = self._get_metadata(opf_xml, 'creator')
            
            nav_points = []
            nav_meta = opf_xml.xpath('//opf:item[contains(@properties, "nav")]', namespaces=self.NAMESPACES)
            if nav_meta:
                nav_href = nav_meta[0].get('href')
                nav_xml = etree.fromstring(zin.read((opf_path.parent / nav_href).as_posix()))
                nav_node = (nav_xml.xpath('//xhtml:nav[@epub:type="toc"]', namespaces=self.NAMESPACES) or 
                            nav_xml.xpath('//xhtml:nav', namespaces=self.NAMESPACES))
                if nav_node and nav_node[0].xpath('./xhtml:ol', namespaces=self.NAMESPACES):
                    nav_points = self._parse_nav_level(nav_node[0].xpath('./xhtml:ol', namespaces=self.NAMESPACES)[0], Path(nav_href).parent)

            ncx_str = f'''<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="auto-gen"/><meta name="dtb:depth" content="3"/></head>
<docTitle><text>{title}</text></docTitle>{"<docAuthor><text>"+author+"</text></docAuthor>" if author else ""}
<navMap>\n{self._build_ncx_points(nav_points)}</navMap></ncx>'''

            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                if 'mimetype' in zin.namelist():
                    zout.writestr('mimetype', zin.read('mimetype'), compress_type=zipfile.ZIP_STORED)
                for item in zin.infolist():
                    if item.filename == 'mimetype': continue
                    data = zin.read(item.filename)
                    if item.filename == opf_path.as_posix():
                        content = data.decode('utf-8').replace('version="3.0"', 'version="2.0"')
                        content = re.sub(r'\s*properties="[^"]*nav[^"]*"', '', content)
                        if 'toc.ncx' not in content:
                            content = content.replace('</manifest>', '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>\n</manifest>')
                            content = re.sub(r'<spine([^>]*)>', r'<spine\1 toc="ncx">', content)
                        zout.writestr(item.filename, content.encode('utf-8'))
                    else:
                        zout.writestr(item, data)
                zout.writestr((opf_path.parent / 'toc.ncx').as_posix(), ncx_str.encode('utf-8'))

# --- 优化后的 GUI 类 ---
class ModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EPUB 3 ➜ 2 转换器")
        self.root.geometry("720x680")
        self.root.configure(bg="#ffffff")
        
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.converter = EpubConverter()
        
        self._setup_ui()

    def _setup_ui(self):
        # 1. 顶部标题 (固定高度)
        header = tk.Frame(self.root, bg="#2c3e50", height=70)
        header.pack(fill="x", side="top")
        tk.Label(header, text="EPUB 3 ➜ 2 转换工具", fg="white", bg="#2c3e50", 
                 font=("Microsoft YaHei", 15, "bold")).pack(pady=20)

        # 主容器：使用 grid 布局来精确控制空间分配
        main_container = tk.Frame(self.root, bg="#ffffff", padx=40, pady=20)
        main_container.pack(fill="both", expand=True)
        
        # 配置列的权重，让中间的内容可以横向拉伸
        main_container.columnconfigure(0, weight=1)

        # 2. 路径选择区 (Grid 第 0, 1 行)
        self._add_path_section(main_container, 0, "源文件目录", self.input_dir, self._select_input)
        self._add_path_section(main_container, 1, "输出目录", self.output_dir, self._select_output)

        # 3. 日志区 (Grid 第 2 行) - 分配权重 weight=1，使其占据剩余所有高度
        log_label = tk.Label(main_container, text="运行日志:", bg="#ffffff", fg="#7f8c8d", 
                             font=("Microsoft YaHei", 9, "bold"))
        log_label.grid(row=2, column=0, sticky="w", pady=(15, 5))
        
        self.log_area = scrolledtext.ScrolledText(main_container, font=("Consolas", 10),
                                                  bg="#f8f9fa", fg="#2c3e50", 
                                                  relief="flat", height=15)
        self.log_area.grid(row=3, column=0, sticky="nsew")
        main_container.rowconfigure(3, weight=1) # 关键：让日志行伸缩

        # 4. 按钮区 (Grid 第 4 行) - 固定在下方，不随窗口伸缩而消失
        self.run_btn = tk.Button(main_container, text="🚀 开始批量转换", bg="#27ae60", fg="white",
                                 font=("Microsoft YaHei", 12, "bold"), relief="flat", 
                                 cursor="hand2", padx=50, pady=12, command=self._start)
        self.run_btn.grid(row=4, column=0, pady=(20, 0))

        # 5. 底部状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief="flat", 
                              anchor="w", bg="#f8f9fa", font=("", 9), padx=10)
        status_bar.pack(side="bottom", fill="x")

    def _add_path_section(self, parent, row_idx, label_text, var, cmd):
        frame = tk.Frame(parent, bg="#ffffff")
        frame.grid(row=row_idx, column=0, sticky="ew", pady=5)
        
        tk.Label(frame, text=label_text, bg="#ffffff", font=("Microsoft YaHei", 9)).pack(anchor="w")
        
        entry_f = tk.Frame(frame, bg="#f8f9fa", padx=8, pady=5)
        entry_f.pack(fill="x", pady=4)
        
        tk.Entry(entry_f, textvariable=var, relief="flat", bg="#f8f9fa", font=("", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(entry_f, text="浏览", bg="#dee2e6", relief="flat", command=cmd, padx=12).pack(side="right")

    def _log(self, msg):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, f" {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def _select_input(self):
        p = filedialog.askdirectory()
        if p: self.input_dir.set(p)

    def _select_output(self):
        p = filedialog.askdirectory()
        if p: self.output_dir.set(p)

    def _start(self):
        if not self.input_dir.get() or not self.output_dir.get():
            messagebox.showwarning("提示", "请先选择目录。")
            return
        self.run_btn.config(state="disabled", text="正在处理...", bg="#95a5a6")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        in_p, out_p = Path(self.input_dir.get()), Path(self.output_dir.get())
        out_p.mkdir(parents=True, exist_ok=True)
        files = list(in_p.glob("*.epub"))
        
        if not files:
            self._log("[!] 未发现 .epub 文件")
        else:
            for f in files:
                try:
                    self.converter.convert(f, out_p / f.name)
                    self._log(f"[✓] 成功: {f.name}")
                except Exception as e:
                    self._log(f"[✗] 失败: {f.name} -> {str(e)}")
            self._log("\n[*] 任务全部完成。")
        
        self.root.after(0, self._on_finish)

    def _on_finish(self):
        self.run_btn.config(state="normal", text="🚀 开始批量转换", bg="#27ae60")
        self.status_var.set("任务已结束")
        messagebox.showinfo("完成", "转换任务已结束。")

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass

    root = tk.Tk()
    app = ModernGUI(root)
    root.mainloop()
