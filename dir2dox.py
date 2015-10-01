#!/usr/bin/python
import os
import argparse
from docx import Document

def log(txt):

    if args.verbose:
        print(txt)

def generate_docx(path, outfile, lvl=1):

    document = Document()

    add_dir_to_dox(path, document, lvl)

    document.save(outfile)

def add_dir_to_dox(dir_path, document, lvl):

    log("Entering %s" % dir_path)

    dir_entries = sorted(os.listdir(dir_path))

    title = os.path.basename(dir_path)

    for entry in dir_entries:
        path = os.path.join(dir_path, entry)

        if os.path.isdir(path):
            document.add_heading(title, lvl)
            add_dir_to_dox(path, document, lvl+1)
        else:
            add_file_to_docx(path, document, lvl)

def add_file_to_docx(path, document, lvl):

    log("Adding %s" % path)

    title = os.path.basename(path)

    document.add_heading(title, lvl)

    with open(path) as f:
        p = document.add_paragraph(f.read())

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-p","--path", type=str,
                    default=".",
                    help="Top level directory path to scan")
    parser.add_argument("-o","--outfile", type=str,
                    default="output.docx",
                    help="Name of output file")
    parser.add_argument("-v", "--verbose", action="store_true",
                    help="verbose")

    args = parser.parse_args()

    generate_docx(args.path, args.outfile)
