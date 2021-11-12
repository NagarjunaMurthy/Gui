import PySimpleGUI as sg
import pandas as pd

sg.theme('dark')

Excel_File='Data_entry.xlsx'
df= pd.read_excel(Excel_File)

layout = [
		[
			sg.Text('Please fill out The Following Details:')
		],
		[
			sg.Text('Name',size=(12,1)),
			sg.InputText(key='Name')
		],
		[	
			sg.Text("Favourite Color",size=(12,1)),
			sg.Combo(['Green','Blue','Red'],key="Favourite Color")
		],
		[
			sg.Text("I Speak",size=(12,1)),
			sg.Checkbox('German',key='German'),
			sg.Checkbox('Spanish',key='Spanish'),
			sg.Checkbox('English',key='English')
		],
		[
			sg.Text('Location',size=(12,1)),
			sg.InputText(key='Location')
		],
		[
			sg.Submit(),
			sg.Button('Clear'),
			sg.Exit()
		]

]

window = sg.Window('Simple data entry Form',layout)

def clear_input():
	for key in values:
		window[key]('')
	return None

while True:
	event,values=window.read()
	if event == sg.WIN_CLOSED or event == 'Exit':
		break
	if event=="Clear":
		clear_input()
	if event == 'Submit':
		df=df.append(values,ignore_index=True)
		df.to_excel(Excel_File,index=False)
		print(event,values)
		sg.popup('Data Saved')
		clear_input()
window.closed()
