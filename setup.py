from cx_Freeze import setup, Executable

setup(
        executables = [Executable("generator.py"), Executable("inkscape.py")]
)