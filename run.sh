#!/bin/bash

# echo "Script executed from: ${PWD}"
# python3 ${PWD}/MtgPriceTracker.py

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$(dirname "$0")"
python3 $DIR/MtgPriceTracker.py
