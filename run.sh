#!/bin/bash
git pull
time python ./extract_CL.py
git add .; git commit -m 'New excel'; git push

