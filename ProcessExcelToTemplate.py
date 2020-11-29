import os
import os.path
import time
import json
import chevron
import pandas
import argparse
from os.path import abspath

def parse_json_recursively(json_object, charToEscape):
	if type(json_object) is dict and json_object:
		for key in json_object:
			if isinstance(json_object[key], str) :
				json_object[key] = json_object[key].replace(charToEscape, "\\"+charToEscape)
	elif type(json_object) is list and json_object:
		for item in json_object:
			parse_json_recursively(item, charToEscape)
			
parser=argparse.ArgumentParser(
    description='''Read excel and use template to generate parsed text.''')
parser.add_argument("-x", "--excel", required=True, help = "Location of excel file")
parser.add_argument("-m", "--mustachetemplate", required=True, help = "Location of template/mustache file")
parser.add_argument("-o", "--output", help = "Location and name of output file")
parser.add_argument("-p", "--escape", help = "Characters to escape")
args=parser.parse_args()

if not os.path.isfile(args.excel):
	print("Excel file not found.")
	exit()
if not os.path.isfile(args.mustachetemplate):
	print("Mustache/Template file not found.")
	exit()


output_dir = ""
output_filename = "Output.txt"
if args.output:
	output_path = abspath(args.output)
	if os.path.isdir(output_path):
		output_dir = output_path
	else:
		if args.output.endswith('/') or args.output.endswith('\\'):
			print("Output path directory does not exist.")
			exit()
		else:
			output_dir = os.path.dirname(output_path) 
			if not os.path.isdir(output_dir):
				print("Parent directory of Output path does not exist.")
				exit()
			output_filename = os.path.basename(output_path)

excel_data_df = pandas.read_excel(args.excel)
json_data = excel_data_df.to_json(orient='records')
json_data = json.loads(json_data)

if args.escape:
	for element in range(0, len(args.escape)):
		parse_json_recursively(json_data, args.escape[element])

with open(args.mustachetemplate, 'r') as template:
	template = template.read()
	template = "{{#JSON_Data_Read}}" + template + "{{/JSON_Data_Read}}"
	
	json_data = {'JSON_Data_Read': json_data}
	renderedText = chevron.render(template, json_data)
	f = open(os.path.join(output_dir, output_filename), 'w')
	f.write(renderedText) 
	f.close()
	print ("Data processed.")
