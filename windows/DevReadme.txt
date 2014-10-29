Windows Development Environment
===============================

Dependencies
-------------

Note: Even if you are developing on a 64-bit system, it is recommended to use a 32-bit tool-chain.

To create a windows development environment you need to install:

- python 2.7
- pywin32
- pyserial
- wxpython
- appdirs
- pywinusb
- mock
- pyhook
- pyinstaller-2.0 (for building for release)

Configuration
-------------

To run Plover, in the root of the repository run `python application\plover`.

You will need to add the root directory of the Plover repository as the PYTHONPATH environment variable. Alternatively, you can temporarily move `application\plover` to the root and change its name to `launch.py` (it may **not** be called `plover.py`) and run it from there as `python launch.py`.

Building
---------

Please ensure you have the pyinstaller-2.0 dependency.

You have to add an empty `__init__.py` file to the `%PYTHON27%\Lib\site-packages\pywinusb` directory.

To build a distributable executable, run `build.bat` with the first argument being the path to `pyinstaller-2.0` (relative to the working directory).
