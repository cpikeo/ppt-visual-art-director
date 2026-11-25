"""打包正式 Skill：仅 SKILL.md、references/、scripts/、requirements*.txt；校验清单、入口引用、解压与 SHA-256。"""
import hashlib, re, sys, zipfile, tempfile, subprocess
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
ver = re.search(r"^version:\s*(\S+)", (ROOT / "SKILL.md").read_text(encoding="utf-8"), re.M).group(1)
files = ["SKILL.md", "requirements.txt", "requirements-render.txt"]
files += sorted(str(p.relative_to(ROOT)) for d in ("references", "scripts") for p in (ROOT / d).rglob("*") if p.is_file() and "__pycache__" not in p.parts and p.suffix != ".pyc")
out = ROOT / "dist"; out.mkdir(exist_ok=True)
name = f"ppt-visual-art-director-{ver}.zip"
z = out / name
if z.exists(): z.unlink()
with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
    for f in files:
        zi = zipfile.ZipInfo(f"ppt-visual-art-director/{f}", (2025, 1, 1, 0, 0, 0)); zi.compress_type = zipfile.ZIP_DEFLATED; zi.external_attr = 0o644 << 16
        zf.writestr(zi, (ROOT / f).read_bytes())
# 核验：清单、无 dev 材料、引用的文件都在包内、解压后可运行
with zipfile.ZipFile(z) as zf:
    assert zf.testzip() is None
    names = zf.namelist()
    assert not any("/dev/" in n or n.endswith(".png") for n in names), "dev 材料混入"
    skill = zf.read("ppt-visual-art-director/SKILL.md").decode("utf-8")
    for ref in set(re.findall(r"`((?:scripts|references)/[\w./-]+)`", skill)):
        assert f"ppt-visual-art-director/{ref}" in names, f"SKILL.md 引用缺失：{ref}"
    with tempfile.TemporaryDirectory() as td:
        zf.extractall(td)
        r = subprocess.run([sys.executable, f"{td}/ppt-visual-art-director/scripts/scaffold.py", "fields", "page"], capture_output=True, text=True)
        assert r.returncode == 0 and "role*" in r.stdout, r.stderr
sha = hashlib.sha256(z.read_bytes()).hexdigest()
(out / "SHA256SUMS").write_text(f"{sha}  {name}\n", encoding="utf-8")
print(name, len(files), "files", z.stat().st_size, "bytes", sha)
