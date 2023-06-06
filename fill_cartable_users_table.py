import psycopg2
import xlrd
import re
sheet_number=int(input('which sheet? : '))
sheet_number-=1
conn = psycopg2.connect(
   database="db_create", user='', password='', host='127.0.0.1', port= '5432'
)
cursor=conn.cursor()
workbook=xlrd.open_workbook("users_list.xlsx")
worksheet = workbook.sheet_by_index(sheet_number)
row_count = worksheet.nrows
for cur_row in range(1, row_count):
	rows = worksheet.row_values(cur_row, start_colx=0, end_colx=None)
	a=rows[0]
	b=rows[1]
	if sheet_number==0:
		cursor.execute("SELECT COUNT(ALL mabda) FROM db_create;")
		res = cursor.fetchall()
		if (res[-1][-1]) < (row_count-1):
		    cursor.execute('INSERT INTO db_create VALUES (%s, %s)', (a, b))
		    conn.commit()

ll=[]
lm=[]
ls=[]
n=[]
maaa=[]
count=0
if sheet_number==1:
	col_mabda = worksheet.col_values(0, start_rowx=1, end_rowx=None)
	col_maghsad = worksheet.col_values(1, start_rowx=1, end_rowx=None)
	r_count = worksheet.nrows
	for i in range(1,r_count):
		r = worksheet.row_values(i, start_colx=0, end_colx=None)
		lm.append(r[0])
		ls.append(r[1])
		n.append(r)
	mylist = list(dict.fromkeys(lm))
	for mab in mylist:
		mc=[]
		tr=[]
		nnn=[]
		cursor.execute("SELECT * FROM db_create WHERE mabda = %s ;",(mab,))
		roo = cursor.fetchall()
		for [mabda, maghsad] in roo:
			if mabda==mab:
				mc.append(maghsad)
				nlist=[]
				for cur_row in range(1, row_count):
					rows = worksheet.row_values(cur_row, start_colx=0, end_colx=None)
					an=rows[0]
					bn=rows[1]
					if an==mab:
						nlist.append(bn)
						tr=list(filter(lambda x: x not in nlist, mc))
			nnn=list(filter(lambda x: x not in mc, nlist))
		for t in tr:
			if len(tr)==0:
				break
			else:
				cursor.execute("DELETE FROM db_create WHERE db_create.mabda = (%s) AND db_create.maghsad = (%s) ;",(mab,t))
				conn.commit()
		for n in nnn:
			if len(nnn)==0:
				break
			else:
				cursor.execute('INSERT INTO db_create VALUES (%s, %s)', (mab, n))
				conn.commit()



	for mabn in mylist:
		l=[]
		for cur_row in range(1, row_count):
			rowsn = worksheet.row_values(cur_row, start_colx=0, end_colx=None)
			ann=rows[0]
			bnn=rows[1]
			mcn=[]
			trn=[]
			nnn=[]
			cursor.execute("SELECT COUNT(ALL mabda) FROM db_create;")
			endres = cursor.fetchall()
			if (endres[-1][-1]) < (row_count-1):
				if ann == mabn:
					l.append([ann,bnn])
					cursor.execute('INSERT INTO db_create VALUES (%s, %s)', (mabn, bnn))
					conn.commit()
					break
		print(l)
				



conn.close()


