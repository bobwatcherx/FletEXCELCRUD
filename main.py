from flet import *
import pandas as pd
# I CREATE RANDOM ID 
import random



def main(page:Page):
	data_dict = []
	# NOW I EXTRACT ALL DATA FROM EXCEL AND CREATE DICT ARRAY
		
	# for remove
	def mydelete(e):
		# GET ID FOR DELETE 
		id_to_delete = e.control.data['id']
		print(id_to_delete)
		data_dict[:] = [data for data in data_dict if data['id'] != id_to_delete]	
		new_df = pd.DataFrame(data_dict)
		new_df.to_excel("datasample.xlsx",index=False)

		# AND REFRESH TABLE
		dt.rows.clear()
		page.update()
		load_data_from_excel()

	# FOR EDIT
	def myedit(e):
		# GET ID FOR EDIT
		id_edit = e.control.data['id']
		name_edit = "YOU EDIT GUYS"
		skor_edit = "YOu SKOR EDIT"
		for data in data_dict:
			if data['id'] == id_edit:
				data['name'] = name_edit
				data['skor'] = skor_edit
		new_df = pd.DataFrame(data_dict)
		new_df.to_excel("datasample.xlsx",index=False)

		# AND REFRESH TABLE
		dt.rows.clear()
		page.update()
		load_data_from_excel()	


	def addnewdata(e):
		# AND I CREATE RANDOM ID FOR ID IN EXCEL
		myid = random.randint(0,100)
		name = con_input.content.controls[0].controls[0].value
		skor = con_input.content.controls[0].controls[1].value
		new_data = {"name":name,"skor":skor,"id":myid}
		data_dict.append(new_data)
		new_df = pd.DataFrame(data_dict)
		new_df.to_excel("datasample.xlsx",index=False)

		# AND REFRESH TABLE
		dt.rows.clear()
		page.update()
		load_data_from_excel()


	def load_data_from_excel():
		df = pd.read_excel("datasample.xlsx")
		# CLEAR data_dict
		data_dict.clear()
		# AND CONVERT TO DICT
		data_dict.extend(df.to_dict(orient="records"))

		# AND NOW LOOP AND ADD TO data_dict
		for x in data_dict:
			dt.rows.append(
				DataRow(
					cells=[
					DataCell(Text(str(x['id']))),
					DataCell(Text(x['name'])),
					DataCell(Text(x['skor'])),
					# AND CREATE 2 button edit and delete
					DataCell(
					Row([
						ElevatedButton("edit",
							data=x,
							on_click=myedit
							),
					ElevatedButton("delete",
					bgcolor="red",color="white",
					# AND GET DATA x
					data=x,
					on_click=mydelete
						),
						])
						)	

					]
					)
				)
		page.update()


	# AND I CREATE CONTAINER INPUT
	con_input = Container(
		content=Column([
			Row([
				TextField(label="name"),
				TextField(label="skor"),
				]),
			ElevatedButton("add new",
				on_click=addnewdata
				)
			])
		)


	# NOW CREATE TABLE
	dt =  DataTable(
		columns=[
			DataColumn(Text("id")),
			DataColumn(Text("name")),
			DataColumn(Text("skor")),
			DataColumn(Text("actions")),
		],
		rows=[]
	)
	load_data_from_excel()

	page.add(
	Column([
		Text("Excel CRUD table",size=30,weight="bold"),
		con_input, 
		dt,
		])
		)

flet.app(target=main)
	