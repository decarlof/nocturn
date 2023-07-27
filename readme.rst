This repository provides python code in support of the interoperability working group part of NOCTURN

Dependencies
============

Create a new conda environment::

	conda create --name nocturn python=3.9

and activate the new environment with::

	conda activate nocturn

then install the following packages::

	conda install xmltodict
	conda install pandas
	conda install openpyxl


Run
===

::
	conda activate nocturn
	python ge.py /nocturn/data/FEG230530_413

ge.py cretes an excel spreasheet called **master.xlsx** containing meta data as defined by the National Museum of Natural History. 

If **master.xlsx** already exists it will append a new meta data row to the existing spreadsheet. In practice you can run ge.py multiple times to automatically populate the excel spreadsheet with meta data::

	python ge.py /nocturn/data/FEG230509_407	
	python ge.py /nocturn/data/FEG230530_408	
	python ge.py /nocturn/data/FEG230530_409	
	python ge.py /nocturn/data/FEG230530_410	
	python ge.py /nocturn/data/FEG230530_411	
	python ge.py /nocturn/data/FEG230530_412	
	python ge.py /nocturn/data/FEG230530_413

will append to the Sheet1 of master.xlsx the meta data for all samples listed above


