#!/bin/bash

container_name="$1"
input="$2"
output="$3"

docker run --rm -v "$(dirname "$input")":/input \
    -v "$(dirname "$output")":/output \
    "$container_name" \
    /input/"$(basename "$input")" \
    /output/"$(basename "$output")"
