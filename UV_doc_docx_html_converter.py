import os
import sys
import re
import csv
import zipfile
import argparse
import docx2txt
import fileinput
from datetime import datetime
import glob
from itertools import islice
import numpy as np
import pandas as pd
import unicodedata 
import xml.etree.ElementTree as ET       #import of diffrent modules needed by the code    
from dateutil.parser import parse
from datetime import datetime
import pypandoc
from tidylib import tidy_document

try:
    import pypandoc
    from tidylib import tidy_document
except ImportError:
    print("\n\nRequires pypandoc and pytidylib. See requirements.txt\n\n")


def convert_to_html(filename):
    # Do the conversion with pandoc
    for filename in glob.glob(os.path.join(filename, '*.doc')):
      output = pypandoc.convert(filename, 'html')

    # Clean up with tidy...
      output, errors = tidy_document(output,  options={
        'numeric-entities': 1,
        'wrap': 80,
    })
      print(errors)

    # replace smart quotes.
      output = output.replace(u"\u2018", '&lsquo;').replace(u"\u2019", '&rsquo;')
      output = output.replace(u"\u201c", "&ldquo;").replace(u"\u201d", "&rdquo;")

    # write the output
      filename, ext = os.path.splitext(filename)
      filename = "{0}.html".format(filename)
      with open(filename, 'w') as f:
        # Python 2 "fix". If this isn't a string, encode it.
        if type(output) is not str:
            output = output.encode('utf-8')
        f.write(output)

      print("Done! Output written to: {}\n".format(filename))


def main():
    if len(sys.argv) == 2:
        convert_to_html(sys.argv[1])
    else:
        print("\nUSAGE: word2html <filename>\n")


if __name__ == "__main__":
    main()
convert_to_html("/home/bioinfo/anakev/wc/UVs_fromblobfiles/more_UVs29092017/") 