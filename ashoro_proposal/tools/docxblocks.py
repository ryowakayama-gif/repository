"""Raw-XML block surgery for word/document.xml.

Splits <w:body> into its direct children as raw XML strings so blocks can be
reordered without ElementTree rewriting namespace prefixes.
"""
import re
import sys

TAG_RE = re.compile(rb"<(/?)([A-Za-z0-9_:.\-]+)((?:[^<>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>", re.S)


def split_body(xml_bytes):
    """Return (head, blocks, tail) where blocks is a list of raw child chunks."""
    m = re.search(rb"<w:body(?:\s[^>]*)?>", xml_bytes)
    if not m:
        raise ValueError("no w:body")
    body_start = m.end()
    close = xml_bytes.rindex(b"</w:body>")
    head = xml_bytes[:body_start]
    tail = xml_bytes[close:]
    inner = xml_bytes[body_start:close]

    blocks = []
    depth = 0
    start = None
    pos = 0
    for t in TAG_RE.finditer(inner):
        closing, name, _attrs, selfclose = t.group(1), t.group(2), t.group(3), t.group(4)
        if name.startswith(b"?") or name.startswith(b"!"):
            continue
        if closing:
            depth -= 1
            if depth == 0:
                blocks.append(inner[start:t.end()])
                start = None
        elif selfclose:
            if depth == 0:
                blocks.append(inner[t.start():t.end()])
        else:
            if depth == 0:
                start = t.start()
                # capture any stray text between blocks (should be none)
            depth += 1
    return head, blocks, tail


def text_of(block):
    """Concatenated w:t / w:delText text of a block."""
    parts = re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", block, re.S)
    s = b"".join(parts).decode("utf-8")
    return (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'"))


def kind(block):
    m = re.match(rb"<(w:[A-Za-z]+)", block)
    return m.group(1).decode() if m else "?"


def assemble(head, blocks, tail):
    return head + b"".join(blocks) + tail


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "unpacked/word/document.xml"
    data = open(path, "rb").read()
    head, blocks, tail = split_body(data)
    for i, b in enumerate(blocks):
        t = text_of(b).replace("\n", "⏎")
        print(f"{i:3d}\t{kind(b):8s}\t{len(b):6d}\t{t[:110]}")
    print(f"# total blocks: {len(blocks)}", file=sys.stderr)
    # round-trip check
    assert assemble(head, blocks, tail) == data, "round-trip mismatch"
    print("# round-trip OK", file=sys.stderr)
