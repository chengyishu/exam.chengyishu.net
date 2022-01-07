#!/bin/bash
cd `dirname $0`
browser-sync start --server --files "**/*.*" --no-notify
# read