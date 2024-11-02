#!/bin/bash

input="$1"
output="$2"

node "$(dirname "$0")"/src/main.js "$input" "$output"
