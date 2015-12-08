#!/bin/sh

. ./clear.sh

git add -u
git add ./doc
git add ./src
git add ./*.sh
git commit -m "Devel"
git push
