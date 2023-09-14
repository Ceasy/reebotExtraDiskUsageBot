from Cython.Build import cythonize
from setuptools import setup, Extension

ext_modules = [Extension("main", ["main.py"])]

setup(
    name="main",
    ext_modules=cythonize(ext_modules)
)
