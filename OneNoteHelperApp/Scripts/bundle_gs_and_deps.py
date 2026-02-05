#!/usr/bin/env python3
"""Bundle Ghostscript (gs) + its Homebrew dylib dependencies into a macOS .app.

Use-case: sandboxed app needs to run gs to convert PS->PDF; Homebrew deps in /opt/homebrew are blocked.

What it does:
- Ensures <App>.app/Contents/MacOS/gs exists (copies from a provided source if given)
- Copies all /opt/homebrew/*.dylib dependencies of gs (recursively) into Contents/Frameworks
- Rewrites install names to use @rpath/<libname>
- Adds @executable_path/../Frameworks rpath to gs
- Fixes libsharpyuv versioned filename by creating libsharpyuv.0.dylib symlink when needed

Then you should codesign the app bundle.

Usage:
  bundle_gs_and_deps.py /path/to/App.app [--gs-src /path/to/gs]
"""

import argparse
import os
import shutil
import subprocess

BREW_PREFIX = "/opt/homebrew"


def run(cmd, check=True):
    return subprocess.run(cmd, check=check, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def otool_deps(path):
    out = subprocess.check_output(["otool", "-L", path], text=True)
    deps = []
    for line in out.splitlines()[1:]:
        line = line.strip()
        if not line:
            continue
        deps.append(line.split(" ", 1)[0])
    return deps


def is_brew_dylib(p: str) -> bool:
    return p.startswith(BREW_PREFIX) and ".dylib" in p


def is_rpath_dylib(p: str) -> bool:
    return p.startswith("@rpath/") and p.endswith(".dylib")


from typing import Optional


def find_brew_dylib_by_basename(base: str) -> Optional[str]:
    """Best-effort lookup for a dylib by basename inside Homebrew.

    We prefer /opt/homebrew/lib, then /opt/homebrew/opt/*/lib.
    """
    direct = os.path.join(BREW_PREFIX, "lib", base)
    if os.path.exists(direct):
        return direct

    # Search in opt/*/lib (limited depth)
    opt_dir = os.path.join(BREW_PREFIX, "opt")
    if os.path.isdir(opt_dir):
        try:
            for formula in os.listdir(opt_dir):
                cand = os.path.join(opt_dir, formula, "lib", base)
                if os.path.exists(cand):
                    return cand
        except Exception:
            pass

    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("app", help="Path to .app")
    ap.add_argument("--gs-src", help="Optional: path to source gs binary to copy into Contents/MacOS/gs")
    args = ap.parse_args()

    app = os.path.abspath(args.app)
    gs_dst = os.path.join(app, "Contents", "MacOS", "gs")
    fw_dir = os.path.join(app, "Contents", "Frameworks")

    os.makedirs(os.path.dirname(gs_dst), exist_ok=True)
    os.makedirs(fw_dir, exist_ok=True)

    if args.gs_src:
        gs_src = os.path.abspath(args.gs_src)
        if not os.path.exists(gs_src):
            raise SystemExit(f"gs-src not found: {gs_src}")
        print("COPY_GS", gs_src, "->", gs_dst)
        shutil.copy2(gs_src, gs_dst)
        os.chmod(gs_dst, 0o755)

    if not os.path.exists(gs_dst):
        raise SystemExit(f"gs not found in app: {gs_dst}")

    # Collect Homebrew dylibs recursively starting from gs.
    # Note: after we rewrite install names to @rpath, `otool -L` will no longer
    # show /opt/homebrew paths. So we do a two-stage approach:
    #  1) gather /opt/homebrew deps from the current binary
    #  2) later, also ensure any @rpath/<lib>.dylib deps are present in Frameworks.
    seen = set()
    queue = [gs_dst]
    brew = set()
    while queue:
        cur = queue.pop(0)
        if cur in seen:
            continue
        seen.add(cur)
        try:
            deps = otool_deps(cur)
        except Exception:
            continue
        for d in deps:
            if is_brew_dylib(d):
                brew.add(d)
                if d not in seen:
                    queue.append(d)

    copied = {}
    for src in sorted(brew):
        base = os.path.basename(src)
        dst = os.path.join(fw_dir, base)
        copied[src] = dst
        if not os.path.exists(dst) or os.path.getsize(dst) != os.path.getsize(src):
            print("COPY", src, "->", dst)
            shutil.copy2(src, dst)
            os.chmod(dst, 0o755)

    # Add rpath to gs.
    print("ADD_RPATH", gs_dst)
    run(["install_name_tool", "-add_rpath", "@executable_path/../Frameworks", gs_dst], check=False)

    def rewrite(path):
        deps = otool_deps(path)
        for d in deps:
            if d in copied:
                new = "@rpath/" + os.path.basename(d)
                print("CHANGE", path, d, "->", new)
                run(["install_name_tool", "-change", d, new, path])

    # Set id for dylibs.
    for src, dst in copied.items():
        base = os.path.basename(src)
        new_id = "@rpath/" + base
        print("ID", dst, "->", new_id)
        run(["install_name_tool", "-id", new_id, dst])

    # Rewrite deps.
    rewrite(gs_dst)
    for dst in copied.values():
        rewrite(dst)

    # Ensure @rpath dylibs referenced by the copied set are actually present.
    # This is important because after rewriting install names, otool will report
    # @rpath/libXYZ.dylib instead of /opt/homebrew/...
    def ensure_rpath_libs(paths_to_scan: list[str]):
        added = True
        while added:
            added = False
            for bin_path in list(paths_to_scan):
                try:
                    deps = otool_deps(bin_path)
                except Exception:
                    continue
                for d in deps:
                    if not is_rpath_dylib(d):
                        continue
                    base = os.path.basename(d)
                    want = os.path.join(fw_dir, base)
                    if os.path.exists(want):
                        continue
                    src = find_brew_dylib_by_basename(base)
                    if not src:
                        continue
                    print("COPY_RPATH", src, "->", want)
                    shutil.copy2(src, want)
                    os.chmod(want, 0o755)
                    # New dylib may itself depend on others.
                    paths_to_scan.append(want)
                    added = True

    scan_list = [gs_dst] + list(copied.values())
    ensure_rpath_libs(scan_list)

    # Special-case: Homebrew webp ships libsharpyuv as versioned file; sometimes deps ask for libsharpyuv.0.dylib.
    sharpyuv_candidates = [
        os.path.join(fw_dir, "libsharpyuv.0.dylib"),
        os.path.join(fw_dir, "libsharpyuv.0.1.2.dylib"),
    ]
    if os.path.exists(sharpyuv_candidates[1]) and not os.path.exists(sharpyuv_candidates[0]):
        print("SYMLINK", sharpyuv_candidates[0], "->", os.path.basename(sharpyuv_candidates[1]))
        os.symlink(os.path.basename(sharpyuv_candidates[1]), sharpyuv_candidates[0])

    # Re-sign the binaries we modified with install_name_tool / copied into Frameworks.
    def adhoc_sign(path: str):
        try:
            run(["/usr/bin/codesign", "--force", "--sign", "-", "--timestamp=none", path], check=True)
            print("CODESIGN", path)
        except Exception as e:
            print("WARN: codesign failed for", path, ":", e)

    adhoc_sign(gs_dst)
    for dst in set(list(copied.values()) + [p for p in scan_list if p.startswith(fw_dir) ]):
        if os.path.exists(dst):
            adhoc_sign(dst)

    print("DONE")


if __name__ == "__main__":
    main()
