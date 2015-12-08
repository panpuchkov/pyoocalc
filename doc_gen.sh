#!/bin/sh

echo "Generation AM4-Server side (python) API documentation"
cd src/
#epydoc --html * -o ../../src/arilot/docs/m/api/
epydoc --html * -o ../doc/tmlibra
cd ..
