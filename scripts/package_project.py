#!/usr/bin/env python3
"""Package Git's nonignored project files without rewriting documentation or evidence."""
import argparse
import hashlib
from pathlib import Path
import subprocess
import zipfile


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=root / '.artifacts/package')
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    listing = subprocess.check_output(
        ['git', 'ls-files', '-z', '--cached', '--others', '--exclude-standard'], cwd=root
    ).decode('utf-8').split('\0')
    names = sorted({n for n in listing if n and (root / n).is_file()
                    and not (root / n).resolve().is_relative_to(output)
                    and not n.endswith(('.bak', '.zip'))})
    if 'ProductMatch.xlsx' not in names:
        parser.error('Run from a prepared Git repository containing ProductMatch.xlsx')
    sums = ''.join(hashlib.sha256((root / n).read_bytes()).hexdigest() + '  ' + n + '\n'
                   for n in names)
    (output / 'SHA256SUMS.txt').write_text(sums, encoding='utf-8')
    archive = output / 'ProductMatch_complete_project.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.write(root / name, 'productmatch/' + name)
        z.writestr('productmatch/SHA256SUMS.txt', sums)
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
    print(f'Packaged {len(names)} project files with checksums: {archive}')


if __name__ == '__main__':
    main()
