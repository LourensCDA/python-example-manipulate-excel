(https://circleci.com/gh/google/pybadges) ![python versions](pyversons.svg)

# Reading EXCEL files using openpyxl

This project was created as an example of using the openpyxl package to read and manipulate an excel file. For more info on openpyxml visit [https://openpyxl.readthedocs.io/en/stable/](https://openpyxl.readthedocs.io/en/stable/).

It is useful for automation to know how to do this manipulation as this allows one to automate repetitive tasks such as importing data to a database.

This repo contains an excel file with some sample data.

## Running

This project uses pipenv.

For more info on installation and usage please visit [https://github.com/pypa/pipenv](https://github.com/pypa/pipenv).

The basics to install pipenv on Windows is

```
pip install pipenv
```

To run the project after installing pipenv:

1.  Open Powershell, Windows Terminal or Command Prompt
2.  Navigate to the location of the repo
3.  Run the following command

```
pipenv run python openpyxl_example.py
```

This will start the virtual enviroment, retrive the relevant packages and run.
