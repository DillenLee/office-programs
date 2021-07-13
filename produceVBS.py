# This code creates the visula basic script to actually do the xml conversion
# It just asks questions regarding the placements of some cells and hard coded
# numbers and creates the appropriate vbs file

import os

title = input('Give this script file a title: ')
creditor = input('Write the creditor code: ')
debtor = input('Write the debtor code: ')
startRowPo = input('Write the beginning row for Po: ')
startRowSo = str(input('Write the beginning row for So (default same as Po): ') or startRowPo)
itemCodeColumnPo = input('Write the column number of the item code for Po: ')
itemCodeColumnSo = str(input('Write the column number of the item code for So (default same as Po): ') or itemCodeColumnPo)
quantityColumnPo = input('Write the column number of the quantity for Po: ')
quantityColumnSo = str(input('Write the column number of the quantity for So (default same as Po): ') or quantityColumnPo)
costColumnPo = input('Write the column number the cost for Po: ')
costColumnSo = str(input('Write the column number the cost for So (default same as Po): ') or costColumnPo)
currencyUnitPo = str(input('Write the currency for Po (default USD): ') or 'USD')
currencyUnitSo = str(input('Write the currency for So (default EUR): ') or 'EUR')


path = os.getcwd()
namePo = path+"/"+title+"Po.vbs"
nameSo = path+"/"+title+"So.vbs"

def code(soOrPo, personCode, row, itemcode, quantity, cost, currency, creditorOrDebtor,name):
        if soOrPo == 'po':
                message ="""

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set filesys = CreateObject("Scripting.FileSystemObject")
Set re = New RegExp
Set xls = CreateObject("Excel.Application")
Set wb = xls.Workbooks.Open(Wscript.Arguments(0))


re.pattern = "\\[^\\]*$"
'path = re.Replace(Wscript.Arguments(0), "")
path = "%s"
Set po = filesys.OpenTextFile(path & "\%s%s.xml", 2, True, -1)

po.writeline("<?xml version=""1.0""?>")
po.writeline("<eExact xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='eExact-Schema.xsd'>")
po.writeline("<Orders>")
po.writeline("<Order type='B'>")
po.writeline("<OrderedAt><%s code='%s'/></OrderedAt>")

set sheet = wb.Sheets(1)

re.pattern = "(.*-.*)-.*"

Row = %s

do
        itemcode = sheet.Cells(row, %s).Value
        quantity = sheet.Cells(row, %s).Value
        price    = sheet.Cells(row, %s).Value

        price    = Replace(Round(price, 3), ",", ".")
        itemcode = re.Replace(itemcode, "$1")

        If Itemcode <> "" Then
                po.writeline("	<OrderLine>")
                po.writeline("		<Item code='" & itemcode & "'/>")
                po.writeline("		<Quantity>" & quantity & "</Quantity>")
                po.writeline("		<Price type='S'><Currency code='%s' /><Value>" & price & "</Value></Price>")
                po.writeline("	</OrderLine>")

                row = row + 1
        End If
loop while Itemcode <> ""

wb.close

po.writeline("</Order>")
po.writeline("</Orders>")
po.writeline("</eExact>")
MsgBox "Created %s%s.xml in " & path

                """%(path,soOrPo,name,creditorOrDebtor ,personCode,row,itemcode,quantity,cost,currency,soOrPo,name)
        else:
                message = """

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set filesys = CreateObject("Scripting.FileSystemObject")
Set re = New RegExp
Set xls = CreateObject("Excel.Application")
Set wb = xls.Workbooks.Open(Wscript.Arguments(0))


re.pattern = "\\[^\\]*$"
'path = re.Replace(Wscript.Arguments(0), "")
path = "%s"
Set po = filesys.OpenTextFile(path & "\%s%s.xml", 2, True, -1)

po.writeline("<?xml version=""1.0""?>")
po.writeline("<eExact xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='eExact-Schema.xsd'>")
po.writeline("<Orders>")
po.writeline("<Order type='V'>")
po.writeline("<YourRef>auto_import</YourRef>")
po.writeline("<OrderedBy><%s code='%s'/></OrderedBy>")

set sheet = wb.Sheets(1)

re.pattern = "(.*-.*)-.*"

Row = %s

do
        itemcode = sheet.Cells(row, %s).Value
        quantity = sheet.Cells(row, %s).Value
        price    = sheet.Cells(row, %s).Value

        price    = Replace(Round(price, 3), ",", ".")
        itemcode = re.Replace(itemcode, "$1")

        If Itemcode <> "" Then
                po.writeline("	<OrderLine>")
                po.writeline("		<Item code='" & itemcode & "'/>")
                po.writeline("		<Quantity>" & quantity & "</Quantity>")
                po.writeline("		<Price type='S'><Currency code='%s' /><Value>" & price & "</Value></Price>")
                po.writeline("	</OrderLine>")

                row = row + 1
        End If
loop while Itemcode <> ""

wb.close

po.writeline("</Order>")
po.writeline("</Orders>")
po.writeline("</eExact>")
MsgBox "Created %s%s.xml in " & path
                """%(path,soOrPo,name,creditorOrDebtor ,personCode,row,itemcode,quantity,cost,currency,soOrPo,name)

        return message


PoScript = open(namePo,'w')
PoScript.write(code('po',creditor,startRowPo,itemCodeColumnPo,quantityColumnPo,costColumnPo,currencyUnitPo,'Creditor',title))
PoScript.close()

SoScript = open(nameSo,'w')
SoScript.write(code('so',debtor,startRowSo,itemCodeColumnSo,quantityColumnSo,costColumnSo,currencyUnitSo,'Debtor',title))
SoScript.close()
