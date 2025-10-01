Attribute VB_Name = "DictSetUp"
Option Explicit


Public Sub ProductDict_SetUp()
	Dim dict As Object
	Set dict = CreateObject("Scripting.Dictionary")

	' Equivalents of Product ID
	dict.Add "Product ID", "Product ID"
	dict.Add "Product Name", "Product ID"
	dict.Add "Product Code", "Product ID"
	dict.Add "Prod. ID", "Product ID"
	dict.Add "SKU", "Product ID"
	dict.Add "Item Number", "Product ID"

	' Equivalents of Region
	dict.Add "North America", "North America"
	dict.Add "N.A.", "North America"
	dict.Add "NORTH AMERICA", "North America"
	dict.Add "north america", "North America"
	dict.Add "Europe", "Europe"
	dict.Add "EUROPE", "Europe"
	dict.Add "europe", "Europe"
	dict.Add "EU", "Europe"
	dict.Add "Europe", "Europe"
	dict.Add "ASIA", "Asia"
	dict.Add "asia", "Asia"
	
	' Equivalents of Quantity sold
	dict.Add "Quantity", "Quantity"
	dict.Add "Qty", "Quantity"
	dict.Add "Qty.", "Quantity"

	' Equivalents of Sales Amount
	dict.Add "Sales Amount", "Sales Amount"
	dict.Add "Sales", "Sales Amount"
	dict.Add "Amount", "Sales Amount"
	
	' Equivalents of Transaction Date
	dict.Add "Transaction Date", "Transaction Date"
	dict.Add "Date", "Transaction Date"
	dict.Add "Trans. Date", "Transaction Date"



	ListDictionarySorted dict, True
End Sub



