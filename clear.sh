#!/bin/sh

PRJ_PATH="./"

echo -n "Removing '*.pyc' files          ":
find "$PRJ_PATH" -iname *.pyc -exec rm {} \; && echo " - OK";
echo -n "Removing '*~' files          ":
find "$PRJ_PATH" -iname *~ -exec rm {} \; && echo " - OK";
