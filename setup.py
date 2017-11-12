from setuptools import setup
from sqlexport_tools import __version__

description = 'SQL Export Tools with a Django Admin Mixin.  Exports CSV and XLSX.'

setup(name='sqlexport-tools',
      packages=['sqlexport_tools'],
      version=__version__,
      description=description,
      author='Matt Buck',
      author_email='matt@mblance.com',
      url='https://github.com/mblance/sqlexport-tools',
      classifiers=[
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: BSD License"
        ],
    )

