#!/usr/bin/python

import mysql.connector
from openpyxl import Workbook
from optparse import OptionParser
import progressbar

parser = OptionParser(conflict_handler='resolve')

parser.add_option("-h", "--host", dest="host",
                  help="DB hostname",default='127.0.0.1')
parser.add_option("-u", "--user", dest="user",
                  help="DB username",default='root')
parser.add_option("-p", "--password", dest="password",
                  help="DB password",default='')
parser.add_option("-d", "--database", dest="database",
                  help="Database name")                  
parser.add_option("-o", "--output", dest="output",
                  help="Output xlsx filename",default='') 
parser.add_option("-v", "--verbose",
                  action="store_true", dest="verbose", default=True,
                  help="Report progress [default]")
parser.add_option("-q", "--quiet",
                  action="store_false", dest="verbose",
                  help="Be quiet")              
                 
(options, args) = parser.parse_args()

if not options.database:
    parser.error('Database name not given')
    
    
options.output = options.output if options.output else options.database + '.xlsx'
                 


cnx = mysql.connector.connect(user=options.user,password=options.password,
                              host=options.host,
                              database=options.database);
query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = %s"
cursor = cnx.cursor()

cursor.execute(query,(options.database,))


tables = [ x for (x,) in cursor]

book = Workbook(write_only=True)


           

for table in tables:
    
    title = table[-31:]
    sheet = book.create_sheet()
    sheet.title = title
    
    query = "SELECT COUNT(*) FROM `%s`" % table
    cursor.execute(query)
    (rows,) = cursor.fetchone()
    
    
    query = "SELECT * FROM `%s`" % table
    cursor.execute(query)

    field_names = [i[0] for i in cursor.description]
    sheet.append(field_names)
    
    print('Exporting %s' % table)
    if options.verbose:
        barCurrent = progressbar.ProgressBar(maxval=rows, widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()]).start()
    
    i=0
    for row in cursor:
        row = [ (x.decode('utf-8') if type(x) is bytearray else x)  for x in row]
        sheet.append(row)
        
        i=i+1
        if options.verbose:
            barCurrent.update(i)
    
    if options.verbose:
        barCurrent.finish()     
    
cursor.close()
if options.verbose:
    print('Writing output file...');

book.save(options.output)


cnx.close()