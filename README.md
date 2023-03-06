# CheBanca! Plugin for [ofxstatement](https://github.com/kedder/ofxstatement/)

Parses CheBanca! xslx statement files to be used with GNU Cash or HomeBank.

## Installation

You can install the plugin as usual from pip or directly from the downloaded git

### `pip`

    pip3 install --user ofxstatement-chebanca

### `setup.py`

    python3 setup.py install --user

## Usage
Download your transactions file from the official bank's site and then run

    ofxstatement convert -t chebanca CheBanca.xlsx CheBanca.ofx


### Loading Historical data

This only supports the statements for the last year, however it's also possible
to convert the old per-quarter statements that are available from the archive as
PDF files.

A plugin is provided that uses `poppler-util`'s `pdftotext` to easily generate
machine parse-able data.

This is an experimental plugin, that may not alwyas work but it can be used via:

    ofxstatement -d convert -t chebanca-pdf ./dir-containing-all-pdfs
    ofxstatement -d convert -t chebanca-pdf CheBanca-20-Q2.pdf
