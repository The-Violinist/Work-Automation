import sqlite3

f = open("urls.txt", "a")
fh = sqlite3.connect("history")
cur = fh.cursor()
## Uncomment to get the correct table name
# for row in cur.execute('select name from sqlite_master where type="table"'):
#     print(row)
for row in cur.execute("select * from urls"):
    f.write(f"{row[1]}\n")
f.close()
fh.close()