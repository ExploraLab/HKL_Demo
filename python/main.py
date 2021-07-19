import sys
import time
import os

import OCR_Core_Main

watchdir = 'data/0_Input'
contents = os.listdir(watchdir)
count = len(watchdir)
dirmtime = os.stat(watchdir).st_mtime

while True:
    newmtime = os.stat(watchdir).st_mtime
    if newmtime != dirmtime:
        dirmtime = newmtime
        newcontents = os.listdir(watchdir)
        added = set(newcontents).difference(contents)
        if added:
            print("Files added: %s" %(" ".join(added)))
            OCR_Core_Main.main()
        removed = set(contents).difference(newcontents)
        if removed:
            print("Files removed: %s" %(" ".join(removed)))

        contents = newcontents
    time.sleep(5)


