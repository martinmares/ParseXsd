#!/bin/bash

FILENAME=oreilly

ruby ../../bin/parsexsd.rb \
    --xsd="$FILENAME.xsd" \
    --stdout
