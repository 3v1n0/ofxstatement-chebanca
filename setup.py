#!/usr/bin/python3
"""Setup
"""
from setuptools import find_packages
from distutils.core import setup

version = "0.2"

with open("README.md") as f:
    long_description = f.read()

setup(
    name="ofxstatement-chebanca",
    version=version,
    author="Marco Trevisan",
    author_email="mail@3v1n0.net",
    url="https://github.com/3v1n0/ofxstatement-chebanca",
    description=("CheBanca! plugin for ofxstatement"),
    long_description=long_description,
    long_description_content_type="text/markdown",
    license="GPLv3",
    keywords=["ofx", "banking", "statement", "chebanca", "che-banca"],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Programming Language :: Python :: 3",
        "Natural Language :: English",
        "Topic :: Office/Business :: Financial :: Accounting",
        "Topic :: Utilities",
        "Environment :: Console",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
    ],
    packages=find_packages("src"),
    package_dir={"": "src"},
    namespace_packages=["ofxstatement", "ofxstatement.plugins"],
    entry_points={
        "ofxstatement": [
            "chebanca = ofxstatement.plugins.chebanca:CheBancaPlugin",
            "chebanca-pdf = ofxstatement.plugins.chebanca_pdf:CheBancaPdfPlugin",
        ],
    },
    install_requires=["ofxstatement"],
    include_package_data=True,
    zip_safe=True,
)
