# Creation of Appendix C

The aim of this package is to automate the creation of Appendix C using the following software: 
- openLCA 
- Brightway2
- SimaPro

The extensions considered in this package are:
- zip (via JSON export to openLCA)
- XLSX (via Brightway2 and Activity Browser)
- CSV (via SimaPro)

## Installation 
  ### Requirements 
      Anaconda installed
      Unzip python package
      
conda env create -f appendix_C.yml

conda activate appendix_C

conda install -c conda-forge tdebonnet/appendix_c
