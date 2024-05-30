from pyhwpx import Hwp
from pyhwpx import *

hwpx = Hwp()

hwp = hwpx.hwp

hwpx.insert_text('hello world')

hwpx.create_table(4,5, True)