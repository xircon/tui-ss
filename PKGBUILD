pkgname=tui-ss-git
pkgver=r1.a44f55a
pkgrel=1
pkgdesc="Modular terminal spreadsheet with a SuperCalc-style slash command workflow"
arch=('any')
url="https://github.com/xircon/tui-ss"
license=('custom:none')
depends=('python')
makedepends=('git')
provides=('tui-ss')
conflicts=('tui-ss')
source=('tui-ss-src::git+https://github.com/xircon/tui-ss.git')
sha256sums=('SKIP')

# Ensure building inside the repo doesn't collide with the cloned source.
srcdir="$startdir/src"

pkgver() {
  cd "$srcdir/tui-ss-src"
  printf "r%s.%s" "$(git rev-list --count HEAD)" "$(git rev-parse --short HEAD)"
}

package() {
  cd "$srcdir/tui-ss-src"

  install -dm755 "$pkgdir/usr/lib/tui-ss"
  cp -r tui_ss "$pkgdir/usr/lib/tui-ss/"
  install -m755 tui-ss "$pkgdir/usr/lib/tui-ss/tui-ss"

  install -dm755 "$pkgdir/usr/bin"
  cat > "$pkgdir/usr/bin/tui-ss" <<'EOF'
#!/usr/bin/env python3
from __future__ import annotations

import sys

ROOT = "/usr/lib/tui-ss"
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tui_ss import main

if __name__ == "__main__":
    raise SystemExit(main())
EOF
  chmod 755 "$pkgdir/usr/bin/tui-ss"

  install -dm755 "$pkgdir/usr/share/doc/tui-ss"
  install -m644 README.md "$pkgdir/usr/share/doc/tui-ss/README.md"
}
