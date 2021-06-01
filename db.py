import MySQLdb
# Open database connection
db = MySQLdb.connect("localhost", "root", "", "rh_v3")

# prepare a cursor object using cursor() method
cursor = db.cursor()

# execute SQL query using execute() method.
#cursor.execute("SELECT VERSION()")

# Fetch a single row using fetchone() method.
#data = cursor.fetchone()
# Use all the SQL you like
cursor.execute("SELECT * FROM Admin")

# print all the first cell of all the rows
for row in cursor.fetchall():
    print(row[0], row[3])

db.close()

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!utility neeed to be verfyed on uapcoming time !!!!!!!!!!!!!!!!!!!!!!
'''
import peewee
from peewee import *

db = MySQLDatabase('jonhydb', user='john', passwd='megajonhy')

class Book(peewee.Model):
    author = peewee.CharField()
    title = peewee.TextField()

    class Meta:
        database = db

Book.create_table()
book = Book(author="me", title='Peewee is cool')
book.save()
for book in Book.filter(author="me"):
    print book.title
'''
