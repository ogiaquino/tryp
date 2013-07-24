try:
    from setuptools import setup
except ImportError:
    from ez_setup import use_setuptools
    use_setuptools()
    from setuptools import setup

requires = ['pandas', 'xlwt']

setup(name='tryp',
      version='0.0',
      description='tryp',
      install_requires=requires,
      tests_require=['nose', ],
      test_suite='nose.collector',
      entry_points={
          'console_scripts': ['trypgen = tryp.tryp:main']
          },
      )
