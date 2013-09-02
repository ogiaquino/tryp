from setuptools import setup

requires = ['pandas', 'xlwt', 'xlrd', 'docopt', 'psycopg2']

setup(name='tryp',
      version='0.1',
      description='tryp',
      long_description="Generate an excel file with crosstabulation defined in tryp file.",
      license='BSD',
      install_requires=requires,
      author='Paulo Patricio Aquino',
      maintainer='Paulo Patricio Aquino',
      maintainer_email='ogiaquino@gmail.com',
      platforms='Linux',
      url='https://github.com/ogiaquino/tryp',
      tests_require=['nose',],
      test_suite='tests',
      entry_points={
          'console_scripts': ['trypgen = tryp.tryp:main']
          },
      packages=['tryp']
      )
