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

cwd = os.path.dirname(os.path.realpath(__file__))
filename = "recipes.xlsx"
outname = os.path.join(cwd, "shopping_list.tex")



#==============================================================================
# Fedest. Autogenerer TeX-kode som kan lave en fin tabel.
#==============================================================================

# Generate TeX code for making the table
texrows = []
data = parse_spreadsheet(filename)
#data = parsed["quantities"]
#details = parsed["descriptions"]
for ingredient, stuff in sorted(data.items(), key = lambda t: t[0]):
    amountdict = stuff["quantities"]
    details = stuff["descriptions"]
    description = ", ".join(details)
    quants = ", ".join([str(round(am, 2))+" "+un for un, am in amountdict.items()])
    cols = [ingredient, quants, description]
    line = u" & ".join(map(str, cols))

    texrows.append(line)

table_contents = "\\\\ \\hline \n".join(texrows)

#==============================================================================
# TeX en indkøbsliste. Også grimt.
#==============================================================================

# Inject into template content
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

tex_code = template.replace("%PLACEHOLDER", table_contents)#.encode("utf-8")

# Compile it!
with open(outname, "w") as f:
    f.write(tex_code)

DEVNULL = open(os.devnull, 'wb')
process = Popen(['pdflatex', outname], stdout=DEVNULL, stderr=DEVNULL)
process.communicate()  # Wait for process to finish

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
