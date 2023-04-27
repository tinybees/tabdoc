#!/usr/bin/env bash
mkdir -p ~/virtual/tabdoc
python3 -m venv ~/virtual/tabdoc
~/virtual/tabdoc/bin/pip install -U pip==20.2.4
~/virtual/tabdoc/bin/pip install wheel
~/virtual/tabdoc/bin/pip install -r requirements.txt
