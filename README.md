# mysql2xlsx
Simple python script for exporting whole mysql database to Microsoft Excel 2007 (xlsx) file. 
The script was optimized for work with large databases containing binary data.

## Installation 
To use the script you need to install the following required modules:

1. openpyxl
2. optparse
3. progressbar
4. mysql.connector

Mysql connector can be downloaded from http://dev.mysql.com/downloads/connector/python/
Other dependencies can be installed using pip
```bash
pip install openpyxl
pip install optparse
pip install progressbar33
```

##Usage
Usage: mysql2xlsx.py [options]

Options:

| Option                             | Description                           |
|------------------------------------|---------------------------------------|
|  --help                            | show this help message and exit       |
|  -h HOST, --host=HOST              | DB hostname                           |
|  -u USER, --user=USER              | DB username                           |
|  -p PASSWORD, --password=PASSWORD  | DB password                           |
|  -d DATABASE, --database=DATABASE  | Database name                         |
|  -o OUTPUT, --output=OUTPUT        | Output xlsx filename                  |
|  -v, --verbose                     | Report progress [default]             |
|  -q, --quiet                       | Be quiet                              |
