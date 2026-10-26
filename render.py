#!/usr/bin/env python3
"""Inject out/data.json into template.html -> out/dashboard.html (self-contained page)."""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))
data = open(os.path.join(HERE, "out", "data.json")).read()
json.loads(data)  # validate
tpl = open(os.path.join(HERE, "template.html")).read()
assert "/*__DATA__*/null" in tpl
html = tpl.replace("/*__DATA__*/null", data.replace("</", "<\\/"))
open(os.path.join(HERE, "out", "dashboard.html"), "w").write(html)
print("wrote out/dashboard.html", len(html), "bytes")
