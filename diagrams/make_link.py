"""Turn a .drawio file into a diagrams.net link that opens the diagram for viewing and editing.

Run:  python make_link.py odpt_platform_layers.drawio
"""
import base64, sys, urllib.parse, zlib

xml = open(sys.argv[1], encoding="utf-8").read()
packer = zlib.compressobj(9, zlib.DEFLATED, -15)  # raw deflate, as diagrams.net expects
packed = packer.compress(urllib.parse.quote(xml, safe="").encode()) + packer.flush()
print("https://app.diagrams.net/#R" + urllib.parse.quote(base64.b64encode(packed).decode(), safe=""))
