import json
import tarfile
import xlsxwriter
import sys
import os
import argparse



def add_to_excel_file(item_src): # Add group members to Excel-file
    workbook = xlsxwriter.Workbook(dict_group_name[item_src]+'.xlsx')
    worksheet = workbook.add_worksheet('Group Members')
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 30)

    row = 0
    
    for item_name, item_address in zip(dict_group[item_src][::2], dict_group[item_src][1::2]):
        worksheet.write(row, 0, item_name)
        worksheet.write(row, 1, item_address)
        row += 1

    workbook.close()

    

def create_dictionary_groups(): # Create dictionary of groups
    for item_obj in obj_json_file:
        if item_obj["type"] == "group":
            dict_group_name[item_obj["uid"]] = item_obj["name"]
            for item_group in item_obj["members"]:
                if item_group["type"] == "network":
                    value.append(item_group["name"])
                    value.append(item_group["subnet4"] + "/" + str(item_group["mask-length4"]))
                if item_group["type"] == "host":
                    value.append(item_group["name"])
                    value.append(item_group["ipv4-address"])
                if item_group["type"] == "simple-gateway":
                    value.append(item_group["name"])
                    value.append(item_group["ipv4-address"])
                if item_group["type"] == "simple-cluster":
                    value.append(item_group["name"])
                    value.append(item_group["ipv4-address"])
                if item_group["type"] == "group":
                    value.append(item_group["name"])
                    value.append(' ')
                
                dict_group[item_obj["uid"]] = dict_group.get(item_obj["uid"], []) + value

                value.clear()

    print("Create group dictionary complete...")
    print()



def local_policy_groups():
    if not os.path.exists(pathname): os.mkdir(pathname) # Make 'a_output' directory if not exist
    os.chdir(pathname) # Directory change to 'a_output'

    create_dictionary_groups() # Create dictionary of groups

    # Open local-policy-file
    network_file_byte = tar.extractfile(network_file).read()
    network_json_file = json.loads(network_file_byte)

    # Parse local-policy-file
    for item_rule in network_json_file:
        if item_rule["type"] == "access-rule":
            for item_one in item_rule["source"]:
                if dict_group.get(item_one, 0) != 0:
                    add_to_excel_file(item_one)
                    del dict_group[item_one]
            for item_one in item_rule["destination"]:
                if dict_group.get(item_one, 0) != 0:
                    add_to_excel_file(item_one)
                    del dict_group[item_one]



def all_policy_groups():
    if not os.path.exists(pathname): os.mkdir(pathname) # Make 'u_output' directory if not exist
    os.chdir(pathname) # Directory change to 'u_output'

    for item_obj in obj_json_file:
        if item_obj["type"] == "group":
            # Create excel-file
            workbook = xlsxwriter.Workbook(item_obj["name"]+'.xlsx')
            worksheet = workbook.add_worksheet('Group Members')
            worksheet.set_column('A:A', 40)
            worksheet.set_column('B:B', 30)

            row = 0

            # Parse object-file   
            for item_group in item_obj["members"]:
                if item_group["type"] == "network":
                    worksheet.write(row, 0, item_group["name"])
                    worksheet.write(row, 1, item_group["subnet4"] + "/" + str(item_group["mask-length4"]))
                if item_group["type"] == "host":
                    worksheet.write(row, 0, item_group["name"])
                    worksheet.write(row, 1, item_group["ipv4-address"])
                if item_group["type"] == "simple-gateway":
                    worksheet.write(row, 0, item_group["name"])
                    worksheet.write(row, 1, item_group["ipv4-address"])
                if item_group["type"] == "simple-cluster":
                    worksheet.write(row, 0, item_group["name"])
                    worksheet.write(row, 1, item_group["ipv4-address"])
                if item_group["type"] == "group":
                    worksheet.write(row, 0, item_group["name"])
                    
                row += 1
                
            workbook.close()



def one_policy_group():
    group_name = input('Input Network Group name: ')
    print()
    row = 0
    exist_group = False

    for item_obj in obj_json_file:
        if item_obj["name"] == group_name:
            exist_group = True
            if item_obj["type"] == "group":
                # Excel-file open
                workbook = xlsxwriter.Workbook(group_name+'.xlsx')
                worksheet = workbook.add_worksheet('Group Members')
                worksheet.set_column('A:A', 40)
                worksheet.set_column('B:B', 30)

                # Excel-file fill by members
                # Parse object-file   
                for item_group in item_obj["members"]:
                    if item_group["type"] == "network":
                        worksheet.write(row, 0, item_group["name"])
                        worksheet.write(row, 1, item_group["subnet4"] + "/" + str(item_group["mask-length4"]))
                    if item_group["type"] == "host":
                        worksheet.write(row, 0, item_group["name"])
                        worksheet.write(row, 1, item_group["ipv4-address"])
                    if item_group["type"] == "simple-gateway":
                        worksheet.write(row, 0, item_group["name"])
                        worksheet.write(row, 1, item_group["ipv4-address"])
                    if item_group["type"] == "simple-cluster":
                        worksheet.write(row, 0, item_group["name"])
                        worksheet.write(row, 1, item_group["ipv4-address"])
                    if item_group["type"] == "group":
                        worksheet.write(row, 0, item_group["name"])
                    row += 1

                workbook.close()
                break
            else:
                print("This name is not Network Group")
                break
    if not exist_group: print("No such group exists\n")
    else: print('Complete\n')



parser = argparse.ArgumentParser(prog='Group members', description='Displays composition groups from Check Point R80 policy to excel-file')
parser.add_argument("gztarfile", help="archive file of FW policy Check Point R80")
parser.add_argument("-a", "--all", help="all groups using in policy", action="store_true")
parser.add_argument("-u", "--ultra", help="all groups from domain", action="store_true")

args = parser.parse_args()

targzfile = args.gztarfile

print('Group members ver. 1.0')
print('https://github.com/Lommaster/Group-Members\n')

dict_group = {}         # Dictionary group by uid
dict_group_name = {}    # Dictionary group by name
value = []              # Value of item group depending network, host, group and etc.

with tarfile.open(targzfile, "r:gz") as tar:
    
    # Index.json load from archive
    index_file_byte = tar.extractfile("index.json").read()
    index_json_file = json.loads(index_file_byte)

    # Name of objects-file 
    obj_file = index_json_file["policyPackages"][0]["objects"]["htmlObjectsFileName"].replace('.html', '.json')

    # Name of local-FW-policy-file
    network_file = index_json_file["policyPackages"][0]["accessLayers"][1]["htmlFileName"].replace('.html', '.json')

    # JSON-file of objects load from archive
    obj_file_byte = tar.extractfile(obj_file).read()
    obj_json_file = json.loads(obj_file_byte)

    if args.all:
        print('Processing...\n')
        pathname = os.path.join(os.getcwd(), 'a_output') # Where will excel-files of groups
        local_policy_groups()
        print('Complete\n')
    elif args.ultra:
        print('Processing...\n')
        pathname = os.path.join(os.getcwd(), 'u_output') # Where will excel-files of groups
        all_policy_groups()
        print('Complete\n')
    else:
        one_policy_group()

tar.close()   
