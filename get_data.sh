#!/bin/sh
# The three Census Annual Retail Trade Survey tables used here (public domain, about 110 KB).
# They're already in data/; this is how to fetch them again or grab a newer year.
set -e
mkdir -p data
base=https://www2.census.gov/programs-surveys/arts/tables
curl -sL -o data/sales.xlsx  "$base/2022benchmarked/sales.xlsx"
curl -sL -o data/invent.xlsx "$base/2022benchmarked/invent.xlsx"
curl -sL -o data/bes.xlsx    "$base/2022/bes.xlsx"
