from setuptools import find_packages, setup

setup(name='rosetta',
      version='0.6',
      description='File-format transformation services',
      author='Tasos Vogiatzoglou',
      author_email='tvoglou@iccode.gr',
      packages=find_packages(),
      scripts=['scripts/rosetta']
 )
