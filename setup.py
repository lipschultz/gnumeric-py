#!/usr/bin/env python

from distutils.core import setup

setup(name='Gnumeric-py',
      version='0.0.1',
      description='A python library for reading and writing Gnumeric files.',
      author='Michael Lipschultz',
      author_email='michael.lipschultz+gpy@gmail.com',
      url='https://github.com/lipschultz/gnumeric-py',
      package_dir={'gnumeric': 'src'},
      packages=['gnumeric'],
      requires=['lxml', 'dateutil'],
      license='GPL3',
      data_files=[('', ['LICENSE', 'README.md'])],
      classifiers=['Development Status :: 3 - Alpha',
                   'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                   'Operating System :: OS Independent',
                   'Programming Language :: Python :: 3',
                   'Topic :: Office/Business :: Financial :: Spreadsheet']
      )
