#!/usr/bin/env python3
"""Create a source-only backup of the *current* Seagate app, no data/venv/binaries.

Usage: python3 backup_current_mac.py '/Volumes/Seagate/App' ~/Desktop/Mitchell_Current.zip
"""
from datetime import datetime
from pathlib import Path
import hashlib
import json
import sys
import zipfile

PATTERNS = (
    'app.py', 'lab_analysis_programs.py', 'mylauncher', 'mylauncher.spec',
    'requirements.txt', 'requirements-*.txt',
    '*.R', '*.Rmd', '*.ipynb',
    'analysis_assets/*.R',
    'cosinor_compound_identifiers.json', 'qba_protocol_profiles.json',
)
EXCLUDE_FRAGMENTS = (
    '_before_', '_backup_', 'backup_', '_tmp_', 'test_',
)


def main(root: Path, destination: Path):
    if not root.is_dir() or not (root/'app.py').is_file():
        raise SystemExit(f'App folder not found: {root}')
    selected=set()
    for pattern in PATTERNS:
        for file in root.glob(pattern):
            if not file.is_file(): continue
            if any(s in file.name.lower() for s in EXCLUDE_FRAGMENTS): continue
            selected.add(file)
    if not selected:
        raise SystemExit('No source files found.')
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.resolve() == root.resolve():
        raise SystemExit('Choose a ZIP file path, not the app directory.')
    if destination.exists():
        raise SystemExit(f'Backup file already exists; choose a new name: {destination}')
    listing=[]
    with zipfile.ZipFile(destination, 'x', compression=zipfile.ZIP_DEFLATED) as z:
        for p in sorted(selected):
            relative=p.relative_to(root).as_posix()
            data=p.read_bytes()
            z.writestr('Mac_Current_Source/'+relative, data)
            listing.append({'file':relative,'size':len(data),'sha256':hashlib.sha256(data).hexdigest()})
        z.writestr('Mac_Current_Source/SOURCE_MANIFEST.json',json.dumps({
            'created':datetime.now().astimezone().isoformat(),
            'source_folder':str(root),
            'files':listing,
        },indent=2))
    print('SUCCESS:',destination)
    print('Source files included:',len(listing))
    print('No user workbooks, R outputs, or Python environments included.')

if __name__=='__main__':
    root=Path(sys.argv[1] if len(sys.argv)>1 else '/Volumes/Seagate/App').expanduser()
    output=Path(sys.argv[2] if len(sys.argv)>2 else '~/Desktop/Mitchell_Current_Source.zip').expanduser()
    main(root,output)
