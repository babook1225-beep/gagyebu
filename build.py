"""Builds dist/ (what the sheet's Loader.gs downloads) from src/."""
import pathlib
root = pathlib.Path(__file__).parent
src, dist = root / 'src', root / 'dist'
dist.mkdir(exist_ok=True)
parser = (src / 'parser.js').read_text()
(dist / 'Parser.js').write_text(parser)
(dist / 'Plan.js').write_text((src / 'plan.js').read_text())
(dist / 'App.js').write_text((src / 'app.js').read_text())
import json, re
plan = (src / 'plan.js').read_text()
body = re.search(r"COLORS: (\{.*?\})", plan, re.S).group(1)
colors = dict(re.findall(r"'([^']+)':\s*'(#[0-9a-fA-F]{6})'", body))
(dist / 'Index.html').write_text((src / 'index.src.html').read_text().replace('/*PARSER*/', parser)
    .replace('/*COLORS*/{}', json.dumps(colors, ensure_ascii=False)))
print(sorted(p.name for p in dist.iterdir()))
