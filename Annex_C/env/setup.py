from setuptools import setup, find_packages
import sm_lib
import os, sys

def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname)).read()

setup(
 
    name='annex_C',
 
    # la version du code
    version=annex_C.1.00,

    packages=find_packages(),

    author="CIRAIG_TD",
 
    author_email="tristan.debonnet@gmail.com",

    description="Creation of annex C",
 
    long_description=open('README.md').read(),
 
    install_requires= ['csv','numpy','pandas','openpyxl','tkinter','pickle'],
 

    url='https://github.com/tdebonnet/Annex_C/',
    classifiers=[]
 
 
)





