#!/bin/bash

# test.sh
#
# This is a hack.
#
# When we use self.model.close(True) to close a spreadsheet file during
# tearDown(), after all the tests end, LibreOffice processes will be
# still left running (confirming by ps auxww | grep libre).
#
# But if we run a dummy libreoffice first before test.py, the
# LibreOffice processes disappear after all the tests end.
#
# However, each time you run test.sh, you have to manually exit
# libreoffice before the test can continue.

if [[ -e ~/.invrcptexporterrc ]]
then
  source ~/.invrcptexporterrc
fi

libreoffice
SCRIPT_PATH=$SCRIPT_PATH $SCRIPT_PATH/test.py
