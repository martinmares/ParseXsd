#!/bin/bash

FILENAME=stackoverflow

ruby ../../bin/parsexsd.rb \
    --xsd="$FILENAME.xsd" \
    --stdout
