Operation =1
Option =0
Begin InputTables
    Name ="tblBooks"
    Name ="qryOrdersSorted"
    Name ="tblOrderDetails"
End
Begin OutputColumns
    Expression ="qryOrdersSorted.LastName"
    Expression ="qryOrdersSorted.FirstName"
    Expression ="qryOrdersSorted.MiddleInit"
    Expression ="qryOrdersSorted.Address"
    Expression ="qryOrdersSorted.City"
    Expression ="qryOrdersSorted.StateOrProvince"
    Expression ="qryOrdersSorted.PostalCode"
    Expression ="qryOrdersSorted.Country"
    Expression ="qryOrdersSorted.PhoneNumber"
    Alias ="CustPayBy"
    Expression ="[qryOrdersSorted].[PayBy]"
    Alias ="CustCCNumber"
    Expression ="[qryOrdersSorted].[CCNumber]"
    Expression ="qryOrdersSorted.OrderID"
    Expression ="qryOrdersSorted.CustomerID"
    Expression ="qryOrdersSorted.StoreID"
    Expression ="qryOrdersSorted.OrderDate"
    Expression ="qryOrdersSorted.ShipName"
    Expression ="qryOrdersSorted.ShipAddress"
    Expression ="qryOrdersSorted.ShipCity"
    Expression ="qryOrdersSorted.ShipStateOrProvince"
    Expression ="qryOrdersSorted.ShipPostalCode"
    Expression ="qryOrdersSorted.ShipCountry"
    Expression ="qryOrdersSorted.ShipPhoneNumber"
    Expression ="qryOrdersSorted.PayBy"
    Expression ="qryOrdersSorted.CCNumber"
    Expression ="qryOrdersSorted.Printed"
    Expression ="qryOrdersSorted.StoreName"
    Expression ="qryOrdersSorted.StAddress"
    Expression ="qryOrdersSorted.StCity"
    Expression ="qryOrdersSorted.StPostalCode"
    Expression ="qryOrdersSorted.StStateOrProvince"
    Expression ="qryOrdersSorted.StCountry"
    Expression ="tblOrderDetails.ISBNNumber"
    Expression ="tblBooks.Title"
    Expression ="tblBooks.EditionNumber"
    Expression ="tblOrderDetails.Quantity"
    Expression ="tblBooks.SuggPrice"
    Expression ="tblOrderDetails.Discount"
End
Begin Joins
    LeftTable ="qryOrdersSorted"
    RightTable ="tblOrderDetails"
    Expression ="[qryOrdersSorted].[OrderID]=[tblOrderDetails].[OrderID]"
    Flag =1
    LeftTable ="tblBooks"
    RightTable ="tblOrderDetails"
    Expression ="[tblBooks].[ISBNNumber]=[tblOrderDetails].[ISBNNumber]"
    Flag =1
End
Begin OrderBy
    Expression ="[qryOrdersSorted].[LastName]"
    Flag =0
    Expression ="[qryOrdersSorted].[FirstName]"
    Flag =0
    Expression ="[qryOrdersSorted].[OrderDate]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =24
    Top =24
    Right =744
    Bottom =408
    Left =-1
    Top =-1
    Right =713
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =344
        Top =4
        Right =440
        Bottom =111
        Top =3
        Name ="tblBooks"
        Name =""
    End
    Begin
        Left =28
        Top =2
        Right =124
        Bottom =109
        Top =9
        Name ="qryOrdersSorted"
        Name =""
    End
    Begin
        Left =200
        Top =2
        Right =296
        Bottom =109
        Top =0
        Name ="tblOrderDetails"
        Name =""
    End
End
