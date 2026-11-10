#!/bin/sh
set -eu

cd "$(dirname "$0")"
rm -rf dist
mkdir dist
cp index.html email-*.html dist/
cp -R img dist/
