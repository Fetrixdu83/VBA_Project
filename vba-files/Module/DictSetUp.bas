Attribute VB_Name = "DictSetUp"
Option Explicit


Public Function ProductDict_SetUp() As Object
	Dim dict As Object
	Set dict = CreateObject("Scripting.Dictionary")
	dict.CompareMode = vbTextCompare

	' Equivalents of Product ID
	dict("Product ID") = "Product ID"
	dict("Product Name") = "Product ID"
	dict("Product Code") = "Product ID"
	dict("Prod. ID") = "Product ID"
	dict("SKU") = "Product ID"
	dict("Item Number") = "Product ID"

	' Equivalent of Region
	dict("Region") = "Region"
	dict("Area") = "Region"
	dict("Zone") = "Region"
	dict("Reg.") = "Region"
	dict("Reg") = "Region"


	' Equivalents of Regions (In line)
	dict("North America") = "North America"
	dict("N.A.") = "North America"
	dict("NORTH AMERICA") = "North America"
	dict("north america") = "North America"
	dict("Europe") = "Europe"
	dict("EUROPE") = "Europe"
	dict("europe") = "Europe"
	dict("EU") = "Europe"
	dict("ASIA") = "Asia"
	dict("asia") = "Asia"
    
	' Equivalents of Quantity sold
	dict("Quantity") = "Quantity Sold"
	dict("Qty") = "Quantity Sold"
	dict("Qty.") = "Quantity Sold"
	dict("Qty. Sold") = "Quantity Sold"
	dict("Qty Sold") = "Quantity Sold"
	dict("Quantity Sold") = "Quantity Sold"

	' Equivalents of Sales Amount
	dict("Sales Amount") = "Sales"
	dict("Sales") = "Sales"
	dict("Amount") = "Sales"
	dict("Revenue") = "Sales"
    
	' Equivalents of Transaction Date
	dict("Transaction Date") = "Transaction Date"
	dict("Date") = "Transaction Date"
	dict("Trans. Date") = "Transaction Date"
	dict("Transaction date") = "Transaction Date"
	dict("transaction date") = "Transaction Date"

	Set ProductDict_SetUp = dict
End Function



