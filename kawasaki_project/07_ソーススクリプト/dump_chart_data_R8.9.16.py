# -*- coding: utf-8 -*-
"""docx 内のグラフ（chartN.xml）の系列値を抜き出し、同名の表の値と突き合わせる。"""
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def charts(path):
    z = zipfile.ZipFile(path)
    out = {}
    for n in sorted(z.namelist()):
        m = re.match(r"word/charts/chart(\d+)\.xml", n)
        if not m:
            continue
        root = ET.fromstring(z.read(n))
        title = "".join(t.text or "" for t in root.iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}t"))
        cats, vals = [], []
        for ser in root.iter("{%s}ser" % NS["c"]):
            cat = ser.find(".//c:cat", NS)
            if cat is not None:
                cats = [p.find("c:v", NS).text for p in cat.iter("{%s}pt" % NS["c"])]
            val = ser.find(".//c:val", NS)
            if val is not None:
                vals.append([p.find("c:v", NS).text
                             for p in val.iter("{%s}pt" % NS["c"])])
        out[int(m.group(1))] = (title, cats, vals)
    return out


if __name__ == "__main__":
  for path in sys.argv[1:]:
      cs = charts(path)
      npts = sum(len(v) for _, _, vs in cs.values() for v in vs)
      print(f"##### {path}: グラフ{len(cs)}点／データ点{npts}")
      for k in sorted(cs):
          t, cats, vals = cs[k]
          flat = [x for v in vals for x in v]
          print(f"  chart{k}: 「{t[:34]}」 系列{len(vals)} 点{len(flat)}"
                f"  {flat[:8]}")
