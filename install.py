#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键安装本仓库技能到本地 AI 技能目录。

仓库中任何直接包含 SKILL.md 的目录视为一个可安装技能。
安装时优先使用 SKILL.md frontmatter 的 name 字段作为技能目录名，
缺省时使用目录名本身。

用法:
    python install.py                          # 安装仓库内全部技能
    python install.py preflight-risk-record    # 只安装指定技能（技能名或目录名）
    python install.py --list                   # 仅列出仓库内可安装的技能
    python install.py --target <目录>          # 指定技能安装目录（默认自动探测）

默认目标目录（按顺序探测第一个存在的目录，均不存在时创建第一个）:
  %LOCALAPPDATA%\\Doubao\\User Data\\Default\\.doubao\\agent_mode\\workspace\\.user_skills
  %USERPROFILE%\\Doubao\\skills
  ./skills

重复安装会覆盖同名技能目录；安装完成后重新打开或新建会话即可生效。
"""
import argparse
import os
import re
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
DEFAULT_TARGET_CANDIDATES = [
    Path(os.environ.get("LOCALAPPDATA", "")) / "Doubao" / "User Data" / "Default"
    / ".doubao" / "agent_mode" / "workspace" / ".user_skills",
    Path(os.environ.get("USERPROFILE", "")) / "Doubao" / "skills",
    Path("skills").resolve(),
]
IGNORE_PATTERNS = shutil.ignore_patterns("__pycache__", "*.pyc", ".git")


def default_target() -> Path:
    for cand in DEFAULT_TARGET_CANDIDATES:
        if cand.is_dir():
            return cand
    target = DEFAULT_TARGET_CANDIDATES[0]
    target.mkdir(parents=True, exist_ok=True)
    return target


def read_skill_name(skill_dir: Path) -> str:
    """从 SKILL.md frontmatter 提取 name；失败时回退为目录名。"""
    try:
        text = skill_dir.joinpath("SKILL.md").read_text(encoding="utf-8")
    except Exception:
        return skill_dir.name
    m = re.match(r"^---\s*\n(.*?)\n---", text, re.DOTALL)
    if not m:
        return skill_dir.name
    fm = m.group(1)
    for idx, line in enumerate(fm.splitlines()):
        mm = re.match(r"^name\s*:\s*(.*)$", line)
        if not mm:
            continue
        v = mm.group(1).strip()
        if v and v not in (">", ">-", "|", "|-", ">+", "|+"):
            return v.strip("'\"")
        # YAML 块标量：取后续第一个非空、非下一个 YAML 键的内容行作为 name
        for nxt in fm.splitlines()[idx + 1:]:
            s = nxt.strip()
            if not s:
                continue
            if re.match(r"^[A-Za-z0-9_\-]+\s*:", s):
                break
            s = s.strip("'\"")
            if s:
                return s
    return skill_dir.name


def find_skills(repo_root: Path):
    """递归查找直接包含 SKILL.md 的目录，返回 (skill_name, skill_dir)。"""
    found = []
    for root, dirs, files in os.walk(repo_root):
        dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", ".venv", "venv", "node_modules")]
        if "SKILL.md" in files:
            p = Path(root)
            found.append((read_skill_name(p), p))
    found.sort(key=lambda x: x[0].lower())
    return found


def main():
    ap = argparse.ArgumentParser(description="一键安装本仓库技能到本地 AI 技能目录")
    ap.add_argument("skill", nargs="?", default=None,
                    help="要安装的技能名（frontmatter name 或目录名）；缺省安装全部")
    ap.add_argument("--target", default=None, help="技能安装目录（默认自动探测豆包个人技能目录）")
    ap.add_argument("--list", action="store_true", help="仅列出仓库内可安装的技能")
    args = ap.parse_args()

    skills = find_skills(REPO_ROOT)
    if not skills:
        print("仓库中未找到任何技能（含 SKILL.md 的目录）。")
        sys.exit(1)

    if args.list:
        print("仓库内可安装的技能：")
        for name, d in skills:
            print(f"  {name:<40s} {d.relative_to(REPO_ROOT)}")
        return

    if args.skill:
        key = args.skill.strip().lower()
        matched = [(n, d) for n, d in skills if n.lower() == key or d.name.lower() == key]
        if not matched:
            print(f"未找到技能: {args.skill}")
            print("可用技能：")
            for name, d in skills:
                print(f"  {name}  ({d.relative_to(REPO_ROOT)})")
            sys.exit(1)
        skills = matched

    target = Path(args.target).resolve() if args.target else default_target()
    target.mkdir(parents=True, exist_ok=True)

    for name, d in skills:
        dst = target / name
        shutil.copytree(d, dst, dirs_exist_ok=True, ignore=IGNORE_PATTERNS)
        n_files = sum(1 for _ in dst.rglob("*") if _.is_file())
        print(f"[installed] {name} -> {dst} ({n_files} files)")

    print(f"\n完成：{len(skills)} 个技能已安装到 {target}")
    print("重新打开或新建会话后，即可在豆包/Agent 中调用这些技能。")


if __name__ == "__main__":
    main()
