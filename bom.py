''' Filtering needed BOM items from file '''

bom_needed = ["SD900.101", "SD900.102", "SD900.104", "SD900.105", "SD900.106", "SD900.107", "SD900.108", "SD900.109", "SD900.110", "SD900.111", "SD900.001", 
    "SD900.003", "SD900.004", "SD900.006", "SD900.008", "SD900.009", "SD900.010", "SD900.011", "SD900.051", "SD900.054", "SD900.056", "SD980.001", "SD980.002", 
    "SD980.005", "SD980.006", "SD980.009", "SD980.120"
]

bom_array = []

with open('bom_items.txt') as my_file:
    for line in my_file:
        for elem in bom_needed:
            if elem in line:
                bom_array.append(line)


f= open("bom_filtered.txt","w+")

for item in bom_array:
     f.write(item)

f.close()