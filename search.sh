#!/bin/bash

read -r -p "What are you looking for: " query

for word in $query; do
    if [ -z ${command+x} ]; then
        command="grep -i '${word}' data/afm.txt"
    else
        command="${command} | grep -i '${word}'"
    fi
done

eval $command
