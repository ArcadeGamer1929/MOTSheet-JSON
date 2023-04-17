#!/usr/bin/env/python3

def show_exception_and_exit(exc_type, exc_value, tb):
	import traceback
	if "root":
		root.update()
		root.destroy()
	traceback.print_exception(exc_type, exc_value, tb)
	input("Press Enter to exit...")
	sys.exit(-1)

import os
import sys
import argparse
import csv
import json
import tkinter as tk
from tkinter import filedialog
if not len(sys.argv) > 1: sys.excepthook = show_exception_and_exit

# MOTSheet-JSON
# version 1.1
# Made by ArcadeGamer1929

print("MOTSheet-JSON v1.1\nType --help for usage\n")
# Usage: motsheet-json.py [-c/--tocsv] [-j/--tojson] [-i/--input] [-o/--output]
parser = argparse.ArgumentParser(description="Converts PD_Tool's MOT JSONs to CSV spreadsheets and vice-versa, useful for copying and moving around different parts of animations."
	+ "\nIt isn't a replacement for MOT Tool, the simplest way to put it is \"dumbed down and complex edit mode\"."
	+ "\nCan be used with or without the command line. LibreOffice Calc is highly recommended for editing and exporting CSV files. Made for PD_Tool (Hatsune Miku: Project DIVA mod tool) by ArcadeGamer1929.",
	formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument("-c", "--tocsv", help="converts a PD_Tool MOT JSON to a CSV spreadsheet", action="store_true")
parser.add_argument("-j", "--tojson", help="converts a CSV spreadsheet to a PD_Tool MOT JSON", action="store_true")
parser.add_argument("-i", "--input")
parser.add_argument("-o", "--output")
args = parser.parse_args()
# jump to line 174 for startup code


def to_intfloat(s):
	try:
		return int(s)
	except ValueError:
		return float(s)

def import_tocsv():
	if inputRead:
		print("reading...", end="", flush=True)

		jsonInput = json.load(inputRead)
		addEmptyCells = 0

		print("\r", end="")
		print("reading... done\n", end="", flush=True)
		print("converting...", end="", flush=True)

		outputWrite.writerow([jsonInput["MOT"][0]["FrameCount"], jsonInput["MOT"][0]["HighBits"]])
		outputWrite.writerow(jsonInput["MOT"][0]["BoneInfo"])

		for i in jsonInput["MOT"][0]["KeySets"]:
			previousFrame = None
			if type(i) != list:
				outputWrite.writerow([])
				previousFrame = None
				continue

			rowArray = [None] * (jsonInput["MOT"][0]["FrameCount"] - 1)
			keysetType = ""
			if i[0] == 1:
				outputWrite.writerow(i[1])
				previousFrame = None
				continue

			for index, j in enumerate(i[1]):
				if i[0] == 2 and index == 0: keysetType = ":"

				if len(j) == 2: j.append("null")
				elif len(j) == 1:
					j.append("null")
					j.append("null")
				if j[0] == previousFrame:
					rowArray[previousFrame] = ("!" if rowArray[previousFrame].find("!") else "") + rowArray[previousFrame] + "/" + "|".join(str(x) for x in j[1:])
				else:
					rowArray.insert(j[0] if len(j) > 1 else 0, "|".join(str(x) for x in j[1:]))
				previousFrame = j[0]

			if addEmptyCells == 1:
				while rowArray[-1] == None:
					rowArray.pop(-1)
			if addEmptyCells != 1: addEmptyCells = 1

			while len(rowArray) > jsonInput["MOT"][0]["FrameCount"]:
				rowArray.pop(-1)
			if type(rowArray[0]) == str: rowArray[0] = keysetType + rowArray[0]
			outputWrite.writerow(rowArray)

		print("\r", end="")
		print("converting... done!\n", end="", flush=True)
		if not len(sys.argv) > 1: input("Press Enter to exit...")

def export_tojson():
	if inputRead:
		print("reading...", end="", flush=True)

		csvInput = csv.reader(inputRead)
		templateJSON = { "MOT": [{ "FrameCount": None, "HighBits": None, "KeySets": [], "BoneInfo": [] }]}
		keySets = templateJSON["MOT"][0]["KeySets"]
		headerRows = -2

		print("\r", end="")
		print("reading... done\n", end="", flush=True)
		print("converting...", end="", flush=True)

		for row in csvInput:
			while len(row) > 0:
				if row[-1] == "":
					row.pop(-1)
				else: break

			if headerRows == -2:
				templateJSON["MOT"][0]["FrameCount"] = int(row[0])
				templateJSON["MOT"][0]["HighBits"] = int(row[1])
				headerRows += 1
				continue
			elif headerRows == -1:
				templateJSON["MOT"][0]["BoneInfo"] = [int(x) for x in row]
				headerRows += 1
				continue

			if len(row) == 0:
				keySets.append(None)
			elif row[0].find(":") == 0:
				row[0] = row[0].replace(":", "")
				keySets.insert(headerRows, [])
				keySets[headerRows].insert(0, 2)
				keySets[headerRows].append([])
			elif len(row) == 1:
				keySets.insert(headerRows, [])
				keySets[headerRows].insert(0, 1)
			elif len(row) > 1:
				keySets.insert(headerRows, [])
				keySets[headerRows].insert(0, 3)
				keySets[headerRows].append([])

			for j, string in enumerate(row):
				if string != "":
					if len(string.split("|")) == 1:
						keySets[headerRows].insert(1, [to_intfloat(string)])
					else:
						if string.find("!") == 0 or 1:
							string = string.replace("!", "")
							for k, splitstring in enumerate(string.split("/")):
								keySets[headerRows][-1].append(splitstring.split("|"))
								keySets[headerRows][-1][-1].pop(0) if splitstring.split("|")[0] == "null" else None
								keySets[headerRows][-1][-1].pop(-1) if splitstring.split("|")[-1] == "null" else None
								for x, y in enumerate(keySets[headerRows][-1][-1]):
									keySets[headerRows][-1][-1][x] = to_intfloat(keySets[headerRows][-1][-1][x])
								keySets[headerRows][-1][-1].insert(0, j)
						else:
							keySets[headerRows][-1].append(string.split("|"))
							keySets[headerRows][-1][-1].pop(0) if string.split("|")[0] == "null" else None
							keySets[headerRows][-1][-1].pop(-1) if string.split("|")[-1] == "null" else None
							for x, y in enumerate(keySets[headerRows][-1][-1]):
								keySets[headerRows][-1][-1][x] = to_intfloat(keySets[headerRows][-1][-1][x])
							keySets[headerRows][-1][-1].insert(0, j)
				# if j < 9 and headerRows == 0: print(keySets[headerRows])
			headerRows += 1
		json.dump(templateJSON, outputWrite)

		print("\r", end="")
		print("converting... done!\n", end="", flush=True)
		if not len(sys.argv) > 1: input("Press Enter to exit...")


saveAsOptions = [[("CSV files", "*.csv"), ("All files", "*.*")], [("JSON files", "*.json"), ("All files", "*.*")]]

if not args.tocsv and not args.tojson:
	print("Choose which conversion to use:")
	print("    >C  convert PD_Tool MOT JSON to CSV")
	print("    >J  convert CSV to PD_Tool MOT JSON")
	x = input()

	if x.lower() == "c": args.tocsv = True; print()
	elif x.lower() == "j": args.tojson = True; print()
	else:
		if not len(sys.argv) > 1:
			print("exiting...")
			input("Press Enter to exit...")
			sys.exit()
		else: sys.exit("exiting...")

root = tk.Tk()
root.withdraw()

if args.tocsv and args.tojson:
	sys.exit("error: both --tocsv and --tojson arguments used, use one or the other")
elif args.tocsv:
	if args.input == None:
		inputRead = open(filedialog.askopenfilename(filetypes = saveAsOptions[args.tocsv or not args.tojson]), "r", newline="")
	else:
		inputRead = open(args.input, "r", newline="")
	if args.output == None:
		outputWrite = csv.writer(open(filedialog.asksaveasfilename(filetypes = saveAsOptions[not args.tocsv or args.tojson], defaultextension="*.*"), "w", newline=""))
	else:
		outputWrite = csv.writer(open(args.output, "w", newline=""))
	root.update()
	root.destroy()
	import_tocsv()

elif args.tojson:
	if args.input == None:
		inputRead = open(filedialog.askopenfilename(filetypes = saveAsOptions[args.tocsv or not args.tojson]), "r", newline="")
	else:
		inputRead = open(args.input, "r", newline="")
	if args.output == None:
		outputWrite = open(filedialog.asksaveasfilename(filetypes = saveAsOptions[not args.tocsv or args.tojson], defaultextension="*.*"), "w", newline="")
	else:
		outputWrite = open(args.output, "w", newline="")
	root.update()
	root.destroy()
	export_tojson()
