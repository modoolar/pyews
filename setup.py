#!/usr/bin/env python
from setuptools import setup

__author__ = ''
__author_email__ = ''
__version__ = '1.0'
__license__ = 'AGPL-v3'
__docs__ = """"""
__url__ = """"""

setup(name='pyews',
      version=__version__,
      description=__docs__,
      author=__author__,
      author_email=__author_email__,
      url=__url__,
      license=__license__,
      platforms=['all'],
      classifiers=[
          'Intended Audience :: Developers',
          'License :: OSI Approved :: AGPL v3',
          'Operating System :: OS Independent',
          'Programming Language :: Python',
          'Programming Language :: Python :: 2',
          'Programming Language :: Python :: 2.5',
          'Programming Language :: Python :: 2.6',
          'Programming Language :: Python :: 2.7',
          'Programming Language :: Python :: 3',
          'Programming Language :: Python :: 3.2',
          'Programming Language :: Python :: 3.3',
          'Programming Language :: Python :: 3.4',
          'Programming Language :: Python :: Implementation :: PyPy',
          'Topic :: Text Processing :: Markup :: XML',
      ],
      py_modules=['pyews'],
      tests_require=['nose>=1.0', 'coverage'],
      install_requires=['requests==2.9.1', 'tornado==4.3'],
      )
