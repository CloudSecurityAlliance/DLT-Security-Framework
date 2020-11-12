#!/usr/bin/env python3

from openpyxl import load_workbook

# Make sure to set data_only=True, we don't want formulas, just data

workbook = load_workbook(filename="Threat Modelling Spreadsheet v2.0.xlsx",  data_only = "True")

#print(workbook.sheetnames)

ws_IRelationships = workbook["IRelationships"]

row_count=0

# mermaid data format:
#```mermaid
#graph LR
#BlockchainAccountSecurityModel(BlockchainAccountSecurityModel)
#CredentialStuffingAttack -->|results in|PasswordTheftAttack[Password Theft Attack]

# Long term read TAGS, split on comma, use that as a dict to hold the strings, then write files, TAG.md

# Initialize some variables here
mermaid_string=""
tags=[]
mermaid_files={}

for row in ws_IRelationships.iter_rows():
    if row_count == 0:
        row_count = row_count + 1
    else:
        # Initialize variables at end of each row, the first row isn't used
        cell_count=1
        for cell in row:
            # A, B, C, D, E, F
            # ID,Type,TAGS,Component/InteractionA,Relationship,Component/InteractionB
            if cell_count == 1:
                # skip ID column
                cell_count = cell_count +1
            elif cell_count == 2:
                # skip Type Column
                cell_count = cell_count +1
            elif cell_count == 3:
                # skip TAGS
                tags=cell.value.split(",")
                # Check all tags and create a dict entry if we don't have one
                for tag_entry in tags:
                    if tag_entry not in mermaid_files:
                        mermaid_files[tag_entry]={}
                        mermaid_files[tag_entry]["name"]=tag_entry
                        mermaid_files[tag_entry]["strings"]=[]
                cell_count = cell_count +1
            elif cell_count == 4:
                # Strip all spaces out for named objecty in mermaid
                mermaid_string=cell.value.replace(" ", "") + " -->|"
                cell_count = cell_count +1
            elif cell_count == 5:
                mermaid_string = mermaid_string + cell.value + "|"
                cell_count = cell_count +1
            elif cell_count == 6:
                # Strip all spaces out for named objecty in mermaid but leave spaces for object title
                mermaid_string = mermaid_string + cell.value.replace(" ", "") + "[" + cell.value + "]"
                cell_count = 1
                mermaid_files[tag_entry]["strings"].append(mermaid_string)
                mermaid_string=""
                tags=[]
print(mermaid_files)
