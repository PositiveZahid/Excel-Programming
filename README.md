# Excel-Programming

Visual Basic Programming in MS Excel

We have some sample data in Excel file. But the data are not well formatted. 

Wrong Format 
-----------------------------------------------------------------
| Product Name		    		|	BP Value	|	Group Name	|
-----------------------------------------------------------------
| Absorbent cotton wool 100g  	|				| 	Woman Care 	|
-----------------------------------------------------------------

Right Format 
-----------------------------------------------------------------
| Product Name		    		|	BP Value	|	Group Name	|
-----------------------------------------------------------------
| Absorbent cotton wool  		|	100g 		| 	Woman Care 	|
-----------------------------------------------------------------

We followed two step to solve this. 

1.	We have to retrieve "100g" from the name of the product 
2.	We have to replace 100g with empty string 

We have written two separate type of function. getMgNumber() and removeMgNumber() 

You can call this function in blank cell and call it as = removeMgNumber(b1), by pressing enter you will get the output result. 

Thanks 
