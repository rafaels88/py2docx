#!/usr/bin/env python


from distutils.core import setup
setup(
    name='py2docx',
    version='0.2.0',
    description='Write .docx documents using Python Code',
    long_description="",
    keywords=['docx', 'word', 'microsoft word'],
    author='Rafael Soares',
    author_email='rafaeltravel88@gmail.com',
    url='http://github.com/rafaels88/py2docx',
    license='MIT',
    packages=['py2docx', 'py2docx.elements', 'py2docx.tmp'],
    include_package_data=True,
    install_requires=['PIL>=1.0']
)
