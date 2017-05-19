#!/bin/bash

if [[ -e ~/.invrcptexporterrc ]]
then
  source ~/.invrcptexporterrc
fi

cd $SCRIPT_PATH

pandoc --from=markdown --to=rst --output=README.rst README.md
rm dist/*
python setup.py sdist
python setup.py bdist_wheel
twine upload dist/*
