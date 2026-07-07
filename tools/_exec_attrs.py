# -*- coding: utf-8 -*-
"""exec 联调：打印 th/td 的真实属性名（不含正则，避免段错误）。只读 HTML。"""
from html.parser import HTMLParser
from pathlib import Path

HTML = Path(r"D:\ai共享盘\MyBrain\tools\_turing_demand.html")


class AttrProbe(HTMLParser):
    def __init__(self):
        super().__init__()
        self.th_attrs = []
        self.td_attrs = []
        self.in_thead = False
        self.stack = []

    def handle_starttag(self, tag, attrs):
        self.stack.append(tag)
        if tag == "thead":
            self.in_thead = True
        if tag == "th" and self.in_thead and len(self.th_attrs) < 4:
            self.th_attrs.append(dict(attrs))
        if tag == "tbody" and len(self.td_attrs) < 4:
            # 只取 tbody 内第一个 tr 的前几个 td：借助 stack 深度判断
            pass
        if tag == "td" and not self.in_thead and len(self.td_attrs) < 4:
            # 仅收集位于 tbody 之后的 td（粗略：遇到 tbody 后）
            self.td_attrs.append(dict(attrs))
        if tag == "tbody":
            self._saw_tbody = True

    def handle_endtag(self, tag):
        if tag == "thead":
            self.in_thead = False
        if self.stack:
            self.stack.pop()


def main():
    html = HTML.read_text(encoding="utf-8")
    p = AttrProbe()
    p.feed(html)
    print("=== TH attrs 样本 ===")
    for a in p.th_attrs:
        print(" ", a)
    print("\n=== TD attrs 样本（tbody 后） ===")
    for a in p.td_attrs:
        print(" ", a)


if __name__ == "__main__":
    main()
