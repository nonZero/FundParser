# FundParser

A simple generic parser for quarterly reports [published][mof] by Israeli funds.

Refer to <https://github.com/nonZero/FundParserTestData> for example data.


# Requirements 
* Python 3.7
* `pandas` and `xlrd`  (See [Pipfile](Pipfile) for more info).

# API

[```fund_parser.extract.extract_fund_data(filename)```](./fund_parser/extract.py) returns a tuple with the following items:

  * `numpy.ndarray|None` data from the "General Data" sheet.
  * `dict[str, pandas.DataFrame|None]` data from all other sheets.

# Demo

`fund2html.py` Creates HTML files from a folder with xlsx files.  To try:

```
git clone --recursive https://github.com/nonZero/FundParser
pip install -r demo_requirements.txt
python fund2html.py test_data out
```

# Development

* Make sure `pipenv` is installed.
* Installation:

        git clone --recursive git@github.com:nonZero/FundParser.git
        pipenv install --dev


* Running tests:

        pipenv run pytest        

[mof]: https://mof.gov.il/hon/Information-entities/Information-and-reports-for-financial-institutions/Pages/Regulation-and-legislation.aspx
