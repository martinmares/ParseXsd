#!/bin/bash

FILENAME=w3cschools

ruby ../../bin/parsexsd.rb \
    --xsd="$FILENAME.xsd" \
    --stdout
