# For pre-compiling the script to boost the startup time a bit.
import py_compile
import os
glpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gamelauncher.py')
py_compile.compile(glpath, glpath + 'c', optimize=2)
input('Press enter to continue...')
