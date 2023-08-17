import xlsxwriter

state = "" # insert state abreviation ex. TX
town = "" # insert town name ex. Austin
allAddresses = []
addresses = []
addressCaseAmt = []

workbook = xlsxwriter.Workbook('CoronaCaseLocations.xlsx')

worksheet = workbook.add_worksheet()

with open('CoronaCaseAddresses.txt', 'r') as casestxt:
	allAddresses = casestxt.read().lower().replace(" street", " st").replace(" road", " rd").replace(" drive", " dr").replace(" cricle", " cir").replace(".", "").replace(",", "").replace("#", "").replace("st1", "st").replace("st2", "st").replace("boro ", "borough ").replace(" marlborough", "").splitlines()

for i in range(len(allAddresses)):
	for j in range(len(allAddresses[i]) - 4):
		if (allAddresses[i][j] == " " and allAddresses[i][j + 1] == "a" and allAddresses[i][j + 2] == "p" and allAddresses[i][j + 3] == "t"):
			newaddress = allAddresses[i][0:j]
			allAddresses[i] = newaddress
			break
		elif (allAddresses[i][j] == "u" and allAddresses[i][j + 1] == "n" and allAddresses[i][j + 2] == "i" and allAddresses[i][j + 3] == "t"):
			newaddress = allAddresses[i][0:j]
			allAddresses[i] = newaddress
			break
		elif (allAddresses[len(allAddresses[i]) - 1] == "0" or allAddresses[len(allAddresses[i]) - 1] == "1" or allAddresses[len(allAddresses[i]) - 1] == "2" or allAddresses[len(allAddresses[i]) - 1] == "3" or allAddresses[len(allAddresses[i]) - 1] == "4" or allAddresses[len(allAddresses[i]) - 1] == "5" or allAddresses[len(allAddresses[i]) - 1] == "6" or allAddresses[len(allAddresses[i]) - 1] == "7" or allAddresses[len(allAddresses[i]) - 1] == "8" or allAddresses[len(allAddresses[i]) - 1] == "9"):
			newaddress = allAddresses[i][0:len(allAddresses[i]) - 2]
			allAddresses[i] = newaddress
			break

for i in range(len(allAddresses)):
	found = False
	for j in range(len(addresses)):
		if addresses[j] == allAddresses[i]:
			addressCaseAmt[j] += 1
			found = True
			break
	if not found:
		addresses.append(allAddresses[i])
		addressCaseAmt.append(1)

worksheet.write('A1', 'Addresses')
for i in range(len(addresses)):
	worksheet.write(i+1, 0, (addresses[i] + " " + town + ", " + state))

worksheet.write('B1', '# Of Clients')
for i in range(len(addressCaseAmt)):
	worksheet.write(i+1, 1, addressCaseAmt[i])

workbook.close()

print(addresses)
print(addressCaseAmt)
