# -*- coding: utf-8 -*-
"""exec 联调解析 v3：用 colid 精确映射 列名->值（只读 HTML 文件）。"""
from html.parser import HTMLParser
from pathlib import Path

HTML = Path(r"D:\ai共享盘\MyBrain\tools\_turing_demand.html")


class VxeParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_thead = False
        self.tbody = None
        self.tbodies = []
        self.row = None
        self.cell = None
        self.colid = None
        self.thead_map = {}      # colid -> 列名（第一个出现）
        self.stack = []

    def handle_starttag(self, tag, attrs):
        self.stack.append(tag)
        d = dict(attrs)
        if tag == "thead":
            self.in_thead = True
        elif tag == "tbody":
            self.tbody = []
            self.tbodies.append(self.tbody)
        elif tag == "tr" and self.stack.count("tr") == 1:
            self.row = {}
        elif tag in ("th", "td"):
            self.cell = []
            self.colid = d.get("colid")

    def handle_endtag(self, tag):
        if tag == "thead":
            self.in_thead = False
        elif tag == "tbody":
            self.tbody = None
        elif tag == "tr" and self.stack.count("tr") == 1:
            if self.tbody is not None and self.row is not None:
                self.tbody.append(self.row)
            self.row = None
        elif tag in ("th", "td"):
            if self.cell is not None and self.colid:
                text = "".join(self.cell).strip()
                if self.in_thead and tag == "th" and self.colid not in self.thead_map:
                    self.thead_map[self.colid] = text
                elif self.row is not None and tag == "td":
                    self.row[self.colid] = text
            self.cell = None
            self.colid = None
        if self.stack:
            self.stack.pop()

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def main():
    html = HTML.read_text(encoding="utf-8")
    p = VxeParser()
    p.feed(html)

    cols = list(p.thead_map.items())   # (colid, 列名)
    print(f"=== 有名列数: {len(cols)}")
    for cid, name in cols:
        print(f"  {cid} = {name}")

    # 定位关键列 colid
    name2cid = {v: k for k, v in p.thead_map.items()}
    want = ["需求单号", "设计款/版型", "需求状态", "领料单状态", "齐套时间"]
    print("\n关键列 colid:", {w: name2cid.get(w) for w in want})

    main_body = max(p.tbodies, key=len) if p.tbodies else []
    print(f"\n=== 主数据行数: {len(main_body)}")
    for i, row in enumerate(main_body[:12]):
        line = []
        for w in want:
            cid = name2cid.get(w)
            if cid:
                line.append(f"{w}={row.get(cid, '')}")
        print(f"  row[{i}]: " + " | ".join(line))


if __name__ == "__main__":
    main()
