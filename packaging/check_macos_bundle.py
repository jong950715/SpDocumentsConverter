"""Reject bundled native libraries that require a newer macOS than we support."""

from pathlib import Path
import plistlib
import sys

from PyInstaller.utils.osx import get_binary_architectures, macosx_version_min


MIN_MACOS_VERSION = '15.5'
MACHO_MAGICS = {
    bytes.fromhex(magic) for magic in (
        'feedface', 'cefaedfe', 'feedfacf', 'cffaedfe',
        'cafebabe', 'bebafeca', 'cafebabf', 'bfbafeca',
    )
}


def check_bundle(bundle):
    bundle = Path(bundle)
    target = tuple(int(part) for part in MIN_MACOS_VERSION.split('.')) + (0,)
    with (bundle / 'Contents/Info.plist').open('rb') as source:
        info = plistlib.load(source)
    errors = []
    if info.get('LSMinimumSystemVersion') != MIN_MACOS_VERSION:
        errors.append(f'Info.plist must declare macOS {MIN_MACOS_VERSION}.')

    checked = 0
    for path in sorted(bundle.rglob('*')):
        if path.is_symlink() or not path.is_file():
            continue
        with path.open('rb') as source:
            if source.read(4) not in MACHO_MAGICS:
                continue
        checked += 1
        _, architectures = get_binary_architectures(str(path))
        relative = path.relative_to(bundle)
        if architectures != ['arm64']:
            errors.append(f'{relative}: expected arm64, found {architectures}.')
            continue
        # Check the finished arm64 slice; a universal binary's Intel slice can
        # advertise an older minimum and hide an incompatible arm64 slice.
        minimum = macosx_version_min(str(path))
        if minimum > target:
            version = '.'.join(str(part) for part in minimum)
            errors.append(f'{relative}: requires macOS {version}.')

    if checked == 0:
        errors.append('No Mach-O binaries found in the application bundle.')
    if errors:
        raise SystemExit(
            f'This bundle does not meet the macOS {MIN_MACOS_VERSION} build requirements:\n'
            + '\n'.join(errors)
            + '\nUse Python, Tcl/Tk and dependencies built for the supported macOS version. '
            'Changing Info.plist or MACOSX_DEPLOYMENT_TARGET cannot repair prebuilt libraries.'
        )
    print(f'Checked {checked} arm64 binaries: minimum OS declarations allow macOS {MIN_MACOS_VERSION}.')


if __name__ == '__main__':
    check_bundle(sys.argv[1])
