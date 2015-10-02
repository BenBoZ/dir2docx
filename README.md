# dir2docx
Python script to generate docx file with contents of directory structure

## Usage

    usage: dir2dox.py [-h] [-a] [-p PATH] [-o OUTFILE] [-v]
    
    optional arguments:
      -h, --help            show this help message and exit
      -a, --all             Parse all files and don't skip over hidden files &
                            folders
      -p PATH, --path PATH  Top level directory path to scan
      -o OUTFILE, --outfile OUTFILE
                            Name of output file
      -v, --verbose         verbose

## Installation

Requires [python-docx](https://github.com/python-openxml/python-docx) module, 
install using pip:

    pip install python-docx

