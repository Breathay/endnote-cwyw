#!/usr/bin/env python3
"""Install this repository as a personal Codex skill without changing settings."""
import argparse
from datetime import datetime, timezone
import os
from pathlib import Path
import shutil


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    codex_home=Path(os.environ.get('CODEX_HOME',Path.home()/'.codex'))
    parser.add_argument('--destination',type=Path,default=codex_home/'skills/endnote-cwyw')
    parser.add_argument('--replace',action='store_true',help='Back up and update an existing skill')
    args=parser.parse_args()
    source=Path(__file__).resolve().parents[1]
    target=args.destination.expanduser().resolve()
    if target==source:
        parser.error('The repository is already at the destination')
    if source in target.parents or target in source.parents:
        parser.error('Source and destination must not contain one another')
    if target.exists() and any(target.iterdir()):
        if not args.replace:
            parser.error('Destination exists; use --replace to back it up before updating')
        stamp=datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
        backup=codex_home/'skill-backups'/('endnote-cwyw-'+stamp)
        shutil.copytree(target,backup,ignore=shutil.ignore_patterns('.venv','__pycache__','*.pyc'))
        print('Backup:',backup)
    target.mkdir(parents=True,exist_ok=True)
    for name in ['SKILL.md','LICENSE','README.md','README.zh-CN.md','pyproject.toml','uv.lock','.python-version','agents','scripts','references','examples']:
        path=source/name
        if path.is_dir():
            shutil.copytree(path,target/name,dirs_exist_ok=True,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
        else:
            shutil.copy2(path,target/name)
    # Older releases installed this file. The backup above retains it when
    # updating; the installed skill now uses pyproject.toml and uv.lock.
    legacy_requirements=target/'requirements.txt'
    if legacy_requirements.is_file():legacy_requirements.unlink()
    print('Installed:',target)


if __name__=='__main__':main()
