# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 15:46:23 2017

@author: omd
"""

from __future__ import division

from collections import defaultdict
import glob
import numpy as np
from worksheet import parse_spreadsheet
import os
import sys
from subprocess import Popen

# Make filename for output files
cwd = os.path.dirname(os.path.realpath(__file__))
cwd = ""

def run(filename):
    basename = os.path.splitext(os.path.basename(filename))[0]
    outname = basename + ".tex"
    txtname = basename + ".txt"

    # Read in data from spreadsheet
    texrows = []
    data = parse_spreadsheet(filename)

    # Format into LaTeX table
    for ingredient, stuff in sorted(data.items(), key = lambda t: t[0]):
        amountdict = stuff["quantities"]
        details = stuff["descriptions"]
        description = ", ".join(details)
        quants = ", ".join([str(round(am, 2))+" "+un for un, am in amountdict.items()])
        cols = [ingredient, quants, description]
        line = u" & ".join(map(str, cols))

        texrows.append(line)
    table_contents = "\\\\ \\hline \n".join(texrows)

    # Make plaintext version as well, in case things go awry
    plaintext = "\n".join(line.replace(" & ", ", ") for line in texrows)
    with open(txtname, "w", encoding="utf-8") as f:
        f.write(plaintext)
    print("Saved plaintext.")

    # LaTeX source to inject table code into
    template = r'''
    \documentclass[a4paper,10pt]{article}

    \usepackage[utf8]{inputenc}
    \usepackage[danish,english]{babel}
    \usepackage[T1]{fontenc}
    \usepackage{graphicx}
    \usepackage{amsmath, amssymb}
    \usepackage{longtable}

    \begin{document}
    \pagestyle{empty}
    \selectlanguage{danish}
    \begin{center}
    \huge Food of doom
    \end{center}
    \begin{longtable}{ccp{8cm}}
    	\hline
         \textbf{Ingredient} & \textbf{Amount} & \textbf{Details} \\ \hline
    	%PLACEHOLDER \\
    	\hline
    \end{longtable}
    \end{document}
    '''

    # Generate a .tex file
    tex_code = template.replace("%PLACEHOLDER", table_contents)
    with open(outname, "w", encoding="utf-8") as f:
        f.write(tex_code)

    # Compile to pdf document
    DEVNULL = open(os.devnull, 'wb')
    process = Popen(['pdflatex', outname], stdout=DEVNULL, stderr=DEVNULL)
    process.communicate()  # Wait for process to finish
    print("Made pdf.")

    # Clean up
    remove_extensions = set(["log", "aux", "gz", "toc", "bbl", "blg", "tex"])
    for ext in remove_extensions:
        target = os.path.join(cwd, outname.replace(".tex", "."+ext))
        try:
            os.remove(target)
        except OSError:
            pass
        #

    print("Job's done!")
    return


if __name__ == '__main__':
    filename = sys.argv[1]
    run(filename)