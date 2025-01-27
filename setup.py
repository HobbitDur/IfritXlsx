from setuptools import setup, find_packages

setup(
    name='IfritXlsx',                 # The package name
    version='1.0.0',                # Version number
    packages=find_packages(),       # Automatically discover all packages and subpackages
    description='Data modifiers for FF8 monsters',  # Short description
    author='hobbitdur',             # Author's name
    url='https://github.com/HobbitDur/IfritXlsx',  # GitHub or project URL
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.12',        # Minimum Python version requirement
)