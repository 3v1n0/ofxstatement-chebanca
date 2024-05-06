# Mediobanca Premier (formerly CheBanca!) Plugin for [ofxstatement](https://github.com/kedder/ofxstatement/)

Parses [Mediobanca Premier](https://mediobancapremier.com) (ex CheBanca!) xslx
statement files to be used with GNU Cash or HomeBank.

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

CheBanca website only allows to download the `xlsx` statements in for the last year,
however it's also possible to get the old statement files in PDF format and convert
these old per-quarter statements that are available from the archive.

A plugin is provided that uses `poppler-util`'s `pdftotext` to easily generate
machine parse-able data.

This is an experimental plugin, that may not always work but it can be used via:

    ofxstatement -d convert -t chebanca-pdf ./dir-containing-all-pdfs
    ofxstatement -d convert -t chebanca-pdf CheBanca-20-Q2.pdf
