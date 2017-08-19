from cx_Freeze import setup,Executable

includefiles = ['replogo.png', 'REPGEN2.png']
includes = []
excludes = []
packages = []

setup(
    name = 'NEW REPORT GEN 1.5',
    version = '1.4',
    description = 'New report gen created for quick generation of report',
    author = 'Renan Rivera',
    author_email = 'renzo031109@gmail.com',
    options = {'build_exe': {'excludes':excludes,'packages':packages,'include_files':includefiles}}, 
    executables = [Executable('newrepgen.py',base = "Win32GUI")]
)
