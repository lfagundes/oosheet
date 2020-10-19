from setuptools import setup, find_packages
import os

setup(name = 'oosheet',
      version = '1.3.1',
      description = 'LibreOffice Spreadsheet scripting library',
      long_description = open(os.path.join(os.path.dirname(__file__), "README")).read(),
      author = "Luis Fagundes",
      author_email = "lhfagundes@gmail.com",
      license = "The MIT License",
      packages = find_packages(),
      entry_points = {
          'console_scripts': [
              'oosheet-pack = oosheet:pack',
              'oosheet-launch = oosheet:launch',
              ]
          },
      # Why isn't install_requires working?
      #install_requires = ['uno'],
      classifiers = [
          'Intended Audience :: Developers',
          'Natural Language :: English',
          'Operating System :: OS Independent',
          'Programming Language :: Python',
          'Topic :: Office/Business :: Financial :: Spreadsheet',
        ],
      url = 'http://oosheet.readthedocs.io/',
      
)
