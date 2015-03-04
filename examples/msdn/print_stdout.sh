#!/bin/bash

FILENAME=msdn

ruby ../../bin/parsexsd.rb \
    --xsd="$FILENAME.xsd" \
    --stdout
