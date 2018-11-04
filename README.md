# Tikkurila-project

Prepation procedure for OpenOffice file generation with openofficegen.py:

1. Install OpenOffice with Py-UNO support opted-in (If OO is already installed,
just run the installator again to add Py-UNO)

2. Copy uno.py and pyuno.pyd files from %OpenOffice installation directory%/program
to %OpenOffice installation directory%/program/python-core-2.7.6/lib

3. You will need to run openofficegen.py using OpenOffice python, not your own -
here's the path: %OO inst. dir%/program/python.exe.