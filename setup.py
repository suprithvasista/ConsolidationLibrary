from setuptools import setup, find_packages
with open ("readme.md", "r") as f:
    description = f.read()
setup(
    name='Pandas-Consolidated-Loader',
    version='0.1',
    packages=find_packages(),
    install_requires=['pandas','xlsxwriter','openpyxl','configparser'],
    long_description=description,
    long_description_content_type="text/markdown",
)