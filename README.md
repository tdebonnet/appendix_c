# Creation of Appendix C

The aim of this package is to automate the creation of Appendix C using the following software: 
- openLCA 
- Brightway2
- SimaPro

The extensions considered in this package are:
- zip (via JSON export to openLCA)
- XLSX (via Brightway2 and Activity Browser)
- CSV (via SimaPro)

## Requirements 
Anaconda installed

Unzip python package

## Installation 

Download and register all python files in appendix_c/Annex_C/annex_c_creation.

conda env create -f appendix_C.yml

conda activate appendix_C

#Select the path where you save the files. 

cd path

python final_interface_appendix_C.py

