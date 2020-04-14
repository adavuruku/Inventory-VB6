Attribute VB_Name = "Module1"
Public old_price
Public old_stock
Public old_code
Public old_location
Public old_name
Public old_date
Public OLDMEASURESD
Public total
Public GOOD
Public GOOD2
Public SGOOD
Public GOOD3
Public GOOD4
Public TYPESS
Public DCLICK
Public sclick
Public dam As Integer
Public DAM2 As Integer
Public COMING
Public nuh
Public fff



Public Sub CONNECTION()
Dim DB As New ADODB.CONNECTION
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=market.mdb;Persist Security Info=False"
End Sub
Public Sub CLOSECONNECTION()
DB.Close
Set DB = Nothing
End Sub

Public Sub clear()
FRMRECEIVE.txtname.Text = ""
FRMRECEIVE.txtprice.Text = ""
FRMRECEIVE.txtstock.Text = ""
FRMRECEIVE.txtlocation.Text = ""
FRMRECEIVE.txtquantity.Text = ""
FRMRECEIVE.txtsupname.Text = ""
FRMRECEIVE.txtsupcomp.Text = ""
FRMRECEIVE.txtsupcontact.Text = ""
FRMRECEIVE.txtamount.Text = ""
FRMRECEIVE.txtcode.Text = ""
FRMRECEIVE.txtgoods.Text = ""
FRMRECEIVE.txtpriceq.Text = ""
FRMRECEIVE.Text1.Text = ""
FRMRECEIVE.txtcode.Locked = True
FRMRECEIVE.txtname.Locked = True
FRMRECEIVE.txtlocation.Locked = True
FRMRECEIVE.txtstock.Locked = True
FRMRECEIVE.txtgoods.Locked = True
FRMRECEIVE.txtpriceq.Locked = True
End Sub

Public Sub uclear()
FRMUPDATE.txtprice.Text = ""
FRMUPDATE.txtcode.Text = ""
FRMUPDATE.txtname.Text = ""
FRMUPDATE.Combo2.Text = ""
FRMUPDATE.txtdate.Text = ""
FRMUPDATE.txtstock.Text = ""
FRMUPDATE.Command2.Enabled = False
FRMUPDATE.Combo3.Text = ""
End Sub
Public Sub sclear()
'FRMSOLD.txtcustname.Text = ""
FRMSOLD.txtquantity.Text = ""
FRMSOLD.txtdiscount.Text = ""
FRMSOLD.txtamount.Text = ""
FRMSOLD.txtcode.Text = ""
FRMSOLD.txtname.Text = ""
FRMSOLD.txtlocation.Text = ""
FRMSOLD.txtprice.Text = ""
FRMSOLD.txtstock.Text = ""

FRMSOLD.txtcustname.Text = FRMSOLD.txtcustname.Text
FRMSOLD.txtcustname.SetFocus
FRMSOLD.txtcustname.SelStart = 0
FRMSOLD.txtcustname.SelLength = Len(FRMSOLD.txtcustname.Text)

FRMSOLD.txtcustname.Locked = True
FRMSOLD.txtquantity.Locked = True
FRMSOLD.txtdiscount.Locked = True
FRMSOLD.txtamount.Locked = True
End Sub
