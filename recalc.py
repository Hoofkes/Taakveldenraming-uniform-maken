import sys, subprocess, os, glob
# Recalculeer met LibreOffice headless via macro (calculateAll) en sla op.
src = sys.argv[1]
macro = '''
import uno
from com.sun.star.beans import PropertyValue
'''
# Gebruik soffice met een basic-macro via commandline is lastig; gebruik python-uno socket.
