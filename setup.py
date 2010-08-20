#!/usr/bin/env python
from setuptools import setup, find_packages

setup(name="zombie",
      version="0.1",
      description="A Python library that converts .docx documents to standards compliant, accessible XHTML.",
      license="MIT",
      author="Jared Lang",
      #author_email="",
      url="http://github.com/kaptainlange/zombie/",
      #dependency_links=[],
      #install_requires=[],
      packages=find_packages(),
      keywords="docx xhtml convert",
      zip_safe= True,)
