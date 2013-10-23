from setuptools import setup

requires = ['pandas', 'xlsxwriter', 'xlrd', 'docopt', 'psycopg2']
long_description = """Generate an excel file with crosstabulation
 defined in tryp file."""

setup(name='tryp',
      version='0.1.1',
      description='tryp',
      long_description=long_description,
      license='BSD',
      install_requires=requires,
      author='Paulo Patricio Aquino',
      author_email='ogiaquino@gmail.com',
      maintainer='Paulo Patricio Aquino',
      maintainer_email='ogiaquino@gmail.com',
      platforms='Linux',
      url='https://github.com/ogiaquino/tryp',
      zip_safe=False,
      tests_require=['nose'],
      test_suite='tests',
      entry_points={
          'console_scripts': ['trypgen = tryp.tryp:main']
          },
      packages=['tryp']
      )
