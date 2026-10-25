"""Synchronize managed skill folders byte-for-byte, preserving backups and local edits."""
from pathlib import Path
import argparse
from datetime import datetime
import hashlib
import json
import re
import shutil
import subprocess
import sys

EXCLUDED = {'__pycache__', '.git'}

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inventory(source):
    manifest = json.loads((source / 'skills-manifest.json').read_text(encoding='utf-8'))
    files = {}
    names = set()
    for skill in manifest['skills']:
        name = skill['name']
        if not re.fullmatch(r'marx-paper(?:-[a-z0-9-]+)?', name) or name in names:
            raise ValueError(f'模块名称无效或重复：{name}')
        names.add(name)
        folder = source / skill['name']
        for required in skill['required_files']:
            if Path(required).is_absolute() or '..' in Path(required).parts:
                raise ValueError(f'资源路径越界：{required}')
            if not (folder / required).is_file():
                raise ValueError(f'源文件缺失：{skill["name"]}/{required}')
        text = (folder / 'SKILL.md').read_text(encoding='utf-8')
        for section in skill.get('required_sections', []):
            if section not in text:
                raise ValueError(f'源工作流不完整：{skill["name"]} 缺少 {section}')
        for path in folder.rglob('*'):
            if path.is_symlink():
                raise ValueError(f'拒绝不明确的符号链接：{path}')
            if path.is_file() and not EXCLUDED.intersection(path.relative_to(folder).parts) and path.suffix != '.pyc':
                files[path.relative_to(source).as_posix()] = digest(path)
    return manifest, files

def git_output(source, *args):
    result = subprocess.run(['git', '-C', str(source), *args], capture_output=True, text=True, encoding='utf-8')
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or 'Git查询失败')
    return result.stdout.strip()

def sync(source, target, apply=False, backup_dir=None, overwrite_installed=False, verify_remote=False):
    source, target = Path(source).resolve(), Path(target).resolve()
    if source == target or source in target.parents or target in source.parents:
        raise ValueError('源和安装目录不可相同或相互包含')
    manifest, expected = inventory(source)
    missing, different = [], []
    for rel, checksum in expected.items():
        destination = target / rel
        if not destination.exists():
            missing.append(rel)
        elif not destination.is_file() or digest(destination) != checksum:
            different.append(rel)
    extra = []
    for skill in manifest['skills']:
        folder = target / skill['name']
        if folder.exists():
            for path in folder.rglob('*'):
                if path.is_file() and not EXCLUDED.intersection(path.relative_to(folder).parts) and path.suffix != '.pyc':
                    rel = path.relative_to(target).as_posix()
                    if rel not in expected:
                        extra.append(rel)
    remote = None
    if verify_remote:
        if git_output(source, 'status', '--porcelain'):
            raise ValueError('源仓库有未提交内容，不能声称GitHub与源一致')
        local = git_output(source, 'rev-parse', 'HEAD')
        remote_lines = git_output(source, 'ls-remote', 'origin', 'refs/heads/main')
        remote = remote_lines.split()[0] if remote_lines else None
        if local != remote:
            raise ValueError(f'GitHub版本不一致：本地 {local}；远程 {remote}')
    if extra:
        raise ValueError('安装目录有源仓库未管理的文件；未删除，请先人工归档：' + ', '.join(extra))
    changes = missing + different
    backup = None
    if apply and changes:
        if not backup_dir:
            raise ValueError('写入前必须指定 --backup-dir')
        state_file = target / '.marx-paper-sync.json'
        previous = json.loads(state_file.read_text(encoding='utf-8')).get('files', {}) if state_file.exists() else {}
        edited = [rel for rel in different if rel not in previous or digest(target / rel) != previous[rel]]
        if edited and not overwrite_installed:
            raise ValueError('安装副本有本地改动或首次同步，需备份后显式 --overwrite-installed：' + ', '.join(edited))
        backup = Path(backup_dir).resolve() / ('codex-before-sync-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f'))
        backup.mkdir(parents=True)
        for rel in different:
            saved = backup / rel
            saved.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(target / rel, saved)
        # Record original hashes so backups can be independently verified.
        (backup / 'backup-manifest.json').write_text(json.dumps({rel: digest(target / rel) for rel in different}, indent=2), encoding='utf-8')
        for rel in changes:
            dest = target / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source / rel, dest)
    if apply:
        for rel, checksum in expected.items():
            if not (target / rel).is_file() or digest(target / rel) != checksum:
                raise RuntimeError(f'同步后核验失败：{rel}')
        target.mkdir(parents=True, exist_ok=True)
        state = {'release': manifest['release'], 'files': expected}
        (target / '.marx-paper-sync.json').write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding='utf-8')
    return {'release': manifest['release'], 'skills': len(manifest['skills']), 'files': len(expected),
            'missing': [] if apply else missing, 'different': [] if apply else different,
            'changed_files': changes if apply else [], 'backup': str(backup) if backup else None,
            'remote_commit': remote, 'consistent': apply or not changes}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--check', action='store_true')
    mode.add_argument('--apply', action='store_true')
    parser.add_argument('--source', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--target', type=Path, default=Path.home() / '.codex/skills')
    parser.add_argument('--backup-dir', type=Path)
    parser.add_argument('--overwrite-installed', action='store_true')
    parser.add_argument('--verify-remote', action='store_true')
    args = parser.parse_args()
    result = sync(args.source, args.target, args.apply, args.backup_dir, args.overwrite_installed, args.verify_remote)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['consistent'] else 1

if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except (ValueError, RuntimeError, OSError) as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(2)
