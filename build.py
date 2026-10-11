"""Builds dist/ (what the sheet's Loader.gs downloads) from src/."""
import pathlib
root = pathlib.Path(__file__).parent
src, dist = root / 'src', root / 'dist'
dist.mkdir(exist_ok=True)
parser = (src / 'parser.js').read_text()
(dist / 'Parser.js').write_text(parser)
(dist / 'Plan.js').write_text((src / 'plan.js').read_text())
(dist / 'App.js').write_text((src / 'app.js').read_text())
(dist / 'Index.html').write_text((src / 'index.src.html').read_text().replace('/*PARSER*/', parser))
print(sorted(p.name for p in dist.iterdir()))
