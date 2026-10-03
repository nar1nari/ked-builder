import subprocess

subprocess.run(
    [
        "uv",
        "run",
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name",
        "ked-builder",
        "--collect-all",
        "ked_builder",
        "src/ked_builder_desktop/__init__.py",
    ],
    check=True,
)
