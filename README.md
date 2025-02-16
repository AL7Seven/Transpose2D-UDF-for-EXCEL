# Transpose2D UDF for EXCEL (or others)

✅ Processing large data sets : Faster and without limitation.

✅ Partial transpose on one or 2 dimensions.

✅ Use with numbers and text : Unlike Application.Transpose, no conversion errors. 

✅ Dynamic table compatible : Seamless integration with new Excel features. 

✅ Use with named ranges and closed workbooks.

The User Defined Function (UDF) AL7_Transpose2D offers an advanced and powerful solution for transposing matrices or ranges of cells into Excel, surpassing Application.Transpose in terms of robustness and flexibility. This function is available in two versions:
•	Transpose2D (alias Trans2D) : 
Primarily intended for use in VBA.
•	TransRange (alias TransRng) : 
Some options are not available in order to be used more easily in the cells. If necessary, use the Transpose2D (Trans2D) function. 

Of course, both functions can be used in cells or in VBA, the difference is that Transpose2D offers more functions than TransRange.
While keeping the advantages of Transpose2D, you can ask to keep the same characteristics/numbering behaviour of transpositions as the Transpose functions of Excel for this put the option LikeExcel = True (-1). 

Additional Features:
PARTIAL TRANSPOSITION of the table indicating the beginning and end of each dimension. 
DEFINE INDEPENDENTLY THE BASIS OF EACH TABLE DIMENSIONS.
