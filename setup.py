from setuptools import setup

description = 'SQL Export Tools with a Django Admin Mixin.  Exports CSV and XLSX.'

setup(name='sqlexport-tools',
      packages=['sqlexport_tools'],
      version='0.0.1',
      description=description,
      author='Matt Buck',
      author_email='matt@mblance.com',
      url='https://github.com/mblance/sqlexport-tools',
      classifiers=[
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: BSD License"
        ],
    )

