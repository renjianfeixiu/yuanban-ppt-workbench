"""Build and smoke-test one release ZIP with the official Node.js runtime."""
import hashlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time
from urllib.request import build_opener, ProxyHandler, urlopen
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


ROOT = "yuanban-ppt-workbench"
NODE = "v22.23.5"
# From https://nodejs.org/dist/v22.23.5/SHASUMS256.txt
RUNTIMES = {
    "macos-arm64": ("darwin-arm64.tar.gz", "e7131995759b61a0170b646010d2ad35a5d37b4cb0d471f3e8317eac6f742f3d"),
    "macos-x64": ("darwin-x64.tar.gz", "c53383238206b9ae4db23c5d8a63645e2e753bec8ab51a49cc81939aa64e9422"),
    "linux-arm64": ("linux-arm64.tar.xz", "2babb9eb5264860d540849f4b064a6fbdebcdc323657d9d382ce23d313fd5e09"),
    "linux-x64": ("linux-x64.tar.xz", "cdeb8670a0c81f51b4dc4ac61919f9b1ad917b304bf8221eb939570663d240f8"),
    "windows-arm64": ("win-arm64.zip", "c1856817f611738c77534efbff8a5a17990593eab24046bc8d2c23a77a160d13"),
    "windows-x64": ("win-x64.zip", "4a191082e9706866830379d577d2cc8a50c776fcd9436a8d061b063ed0d9a0e3"),
}


def read_runtime(target, local_archive):
    suffix, checksum = RUNTIMES[target]
    filename = f"node-{NODE}-{suffix}"
    if local_archive:
        data = Path(local_archive).read_bytes()
    else:
        with urlopen(f"https://nodejs.org/dist/{NODE}/{filename}", timeout=90) as response:
            data = response.read()
    if hashlib.sha256(data).hexdigest() != checksum:
        raise ValueError(f"Node checksum mismatch for {filename}")
    prefix = filename.split(".tar.")[0].removesuffix(".zip")
    if suffix.endswith(".zip"):
        with ZipFile(io.BytesIO(data)) as archive:
            return archive.read(f"{prefix}/node.exe"), archive.read(f"{prefix}/LICENSE")
    with tarfile.open(fileobj=io.BytesIO(data)) as archive:
        return archive.extractfile(f"{prefix}/bin/node").read(), archive.extractfile(f"{prefix}/LICENSE").read()


def add_file(archive, name, data, mode):
    item = ZipInfo(f"{ROOT}/{name}")
    item.create_system = 3
    item.external_attr = (0o100000 | mode) << 16
    archive.writestr(item, data, compress_type=ZIP_DEFLATED, compresslevel=5)


def build(base, target, tag, local_archive=None):
    binary, license_text = read_runtime(target, local_archive)
    windows = target.startswith("windows-")
    node = "node.exe" if windows else "node"
    if windows:
        launcher = "START.cmd"
        command = '@echo off\r\ncd /d "%~dp0"\r\n"runtime\\node.exe" "product-demo\\server.mjs" --open-browser\r\nif errorlevel 1 pause\r\n'
    else:
        launcher = "START.command" if target.startswith("macos-") else "START.sh"
        command = '#!/bin/sh\nset -eu\ncd -P "$(dirname "$0")"\nexec ./runtime/node product-demo/server.mjs --open-browser\n'
    out = Path("dist") / f"{ROOT}-{tag}-{target}.zip"
    out.parent.mkdir(exist_ok=True)
    with ZipFile(base) as source, ZipFile(out, "x") as output:
        for item in source.infolist():
            data = source.read(item)
            if item.filename == f"{ROOT}/README.md":
                intro = f"# 免安装启动\n\n解压整个文件夹，打开 `{launcher}`，浏览器将进入工作台。这个包已包含 Node.js {NODE} 和运行依赖。\n\n"
                data = intro.encode("utf-8") + data
            output.writestr(item, data)
        add_file(output, f"runtime/{node}", binary, 0o755)
        add_file(output, "runtime/LICENSE", license_text, 0o644)
        add_file(output, launcher, command.encode("ascii"), 0o755)
    return out, node, launcher


def check_launcher(root, launcher, env, temporary):
    windows = os.name == "nt"
    cmd = [os.environ["COMSPEC"], "/d", "/c", launcher] if windows else [str(root / launcher)]
    log = temporary / "startup.log"
    opener = build_opener(ProxyHandler({}))
    with log.open("wb") as stream:
        child = subprocess.Popen(cmd, cwd=root, env=env, stdout=stream, stderr=stream)
        try:
            for _ in range(120):
                output = log.read_bytes()
                url = re.search(rb"http://127\.0\.0\.1:\d+", output)
                if url:
                    with opener.open(url.group().decode() + "/api/health", timeout=5) as response:
                        assert json.load(response)["status"] == "ready"
                    return
                if child.poll() is not None:
                    break
                time.sleep(0.1)
            raise RuntimeError(f"Launcher did not start the workbench: {log.read_bytes()!r}")
        finally:
            if child.poll() is None:
                if windows:
                    subprocess.run(["taskkill", "/PID", str(child.pid), "/T", "/F"], check=True)
                else:
                    child.terminate()
                child.wait(timeout=15)


def verify(output, node, launcher, target):
    with tempfile.TemporaryDirectory() as directory:
        temporary = Path(directory)
        root = temporary / ROOT
        with ZipFile(output) as archive:
            archive.extractall(temporary)
            for name in [launcher, f"runtime/{node}"]:
                mode = archive.getinfo(f"{ROOT}/{name}").external_attr >> 16
                assert mode & 0o111
                (root / name).chmod(mode & 0o777)
        env = {**os.environ, "PPT_DEMO_NO_BROWSER": "1", "PPT_DEMO_PORT": "0",
               "PPT_DEMO_DATA_DIR": str(temporary / "project")}
        if os.name == "nt":
            env["PATH"] = str(Path(os.environ["SystemRoot"]) / "System32")
        else:
            commands = temporary / "system-commands"
            commands.mkdir()
            (commands / "dirname").symlink_to(shutil.which("dirname"))
            env["PATH"] = str(commands)
        env.pop("NODE_PATH", None)
        env.pop("NODE_OPTIONS", None)
        exe = root / "runtime" / node
        identity = subprocess.check_output([exe, "-p", "process.version + '/' + process.arch"], env=env, text=True)
        assert identity.strip() == NODE + "/" + target.rsplit("-", 1)[1], identity
        subprocess.run([exe, "--test", "product-demo/tests/workbench.test.mjs"], cwd=root, env=env, check=True, timeout=90)
        check_launcher(root, launcher, env, temporary)
    print(f"Verified {output.name}, SHA-256 {hashlib.sha256(output.read_bytes()).hexdigest()}", flush=True)


if __name__ == "__main__":
    if len(sys.argv) not in (4, 5):
        raise SystemExit("Usage: package-native.py SOURCE_ZIP PLATFORM vX.Y.Z [NODE_ARCHIVE]")
    base, target, tag, *local_archive = sys.argv[1:]
    if not re.fullmatch(r"v\d+\.\d+\.\d+", tag) or target not in RUNTIMES:
        raise SystemExit("Expected a release tag and a supported platform")
    verify(*build(Path(base), target, tag, *local_archive), target)
