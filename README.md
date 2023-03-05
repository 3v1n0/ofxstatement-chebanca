# CheBanca! Plugin for [ofxstatement](https://github.com/kedder/ofxstatement/)

Parses CheBanca! xslx statement files to be used with GNU Cash or HomeBank.

## Installation

You can install the plugin as usual from pip or directly from the downloaded git

### `pip`

    pip3 install --user ofxstatement-intesasp

### `setup.py`

    python3 setup.py install --user

## Configuration

To edit the config file run this command

    ofxstatement edit-config

Save and exit the text editor

## Usage
Download your transactions file from the official bank's site and then run

    ofxstatement convert -t chebanca CheBanca.xlsx CheBanca.ofx
