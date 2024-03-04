Imports System.Data.SqlClient

Public Class PaymentCRUD

    Private db As New DatabaseHelper()

    Public Function InsertOrder(order As ModelLibrary.BookOrder) As Boolean
        'dateTimePayment As DateTime, userID As Integer, contact As String, totalPrice As Decimal,
        'method As String, deliveryFee As Decimal, location As String
        Dim dateTimePayment As DateTime = order.paymentDateTime
        Dim userID As Integer = order.userID
        Dim contact As String = order.orderContact
        Dim totalPrice As Decimal = order.paymentAmount
        Dim method As String = order.orderDeliveryMethod
        Dim deliveryFee As Decimal = order.orderDeliveryFee
        Dim location As String = order.orderLocation
        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then

                'update order table
                strSql = "INSERT into BookOrder(orderDateTime, userID, orderContact, orderSubtotal," +
                "orderDeliveryMethod, orderDeliveryFee, orderLocation) VALUES (@orderDateTime, @userID," +
                "@orderContact, @orderSubtotal, @orderDeliveryMethod, @orderDeliveryFee, @orderLocation)"

                MySqlCommand = New SqlCommand(strSql, db.conn)

                MySqlCommand.Parameters.AddWithValue("@orderDateTime", dateTimePayment)
                MySqlCommand.Parameters.AddWithValue("@userID", userID)
                MySqlCommand.Parameters.AddWithValue("@orderContact", contact)
                MySqlCommand.Parameters.AddWithValue("@orderSubtotal", totalPrice)
                MySqlCommand.Parameters.AddWithValue("@orderDeliveryMethod", method)
                MySqlCommand.Parameters.AddWithValue("@orderDeliveryFee", deliveryFee)
                MySqlCommand.Parameters.AddWithValue("@orderLocation", location)
                MySqlCommand.ExecuteNonQuery()

                db.CloseConnection()

            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function
    

    Public Function InsertPayment(payment As ModelLibrary.PaymentDetailViewModel) As Boolean
        'dateTimePayment As DateTime, userID As Integer, paymentAmount As Decimal,
        'country As String, fname As String, lname As String, address As String,
        'postcode As String, city As String, phone As String
        Dim dateTimePayment As DateTime = payment.dateTimePayment
        Dim userID As Integer = payment.userID
        Dim paymentAmount As Decimal = payment.totalPrice
        Dim country As String = payment.country
        Dim fname As String = payment.fname
        Dim lname As String = payment.lname
        Dim address As String = payment.address
        Dim postcode As String = payment.postcode
        Dim city As String = payment.city
        Dim phone As String = payment.phone

        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then

                strSql = "Select orderID from BookOrder Where orderDateTime = @orderDateTime And userID = @userID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@orderDateTime", dateTimePayment)
                MySqlCommand.Parameters.AddWithValue("@userID", userID)
                Dim orderID = MySqlCommand.ExecuteScalar()

                strSql = "INSERT into Payment(paymentAmount, paymentDateTime, paymentCountry, paymentFname," +
                "paymentLname, paymentAddress, paymentPostcode, paymentCity, paymentPhone, orderID) VALUES (@paymentAmount, @paymentDateTime," +
                "@paymentCountry, @paymentFname, @paymentLname, @paymentAddress," +
                "@paymentPostcode, @paymentCity, @paymentPhone, @orderID)"

                MySqlCommand = New SqlCommand(strSql, db.conn)

                MySqlCommand.Parameters.AddWithValue("@paymentAmount", paymentAmount)
                MySqlCommand.Parameters.AddWithValue("@paymentDateTime", dateTimePayment)
                MySqlCommand.Parameters.AddWithValue("@paymentCountry", country)
                MySqlCommand.Parameters.AddWithValue("@paymentFname", fname)
                MySqlCommand.Parameters.AddWithValue("@paymentLname", lname)
                MySqlCommand.Parameters.AddWithValue("@paymentAddress", address)
                MySqlCommand.Parameters.AddWithValue("@paymentPostcode", postcode)
                MySqlCommand.Parameters.AddWithValue("@paymentCity", city)
                MySqlCommand.Parameters.AddWithValue("@paymentPhone", phone)
                MySqlCommand.Parameters.AddWithValue("@orderID", orderID)

                MySqlCommand.ExecuteNonQuery()

                db.CloseConnection()
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Public Function UpdateBookATC(orderDetail As ModelLibrary.OrderDetail) As Boolean
        'qsBookID As Integer, qsQuantity As Integer, dateTimePayment As DateTime,
        '                          userID As Integer, paymentAmount As Decimal,
        '                       country As String, fname As String, lname As String, address As String,
        '                       postcode As String, city As String, phone As String
        Dim qsBookID As Integer = orderDetail.bookID
        Dim qsQuantity As Integer = orderDetail.odQuantity
        Dim dateTimePayment As DateTime = orderDetail.orderDateTime
        Dim userID As Integer = orderDetail.userID
        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then
                strSql = "Select orderID from BookOrder Where orderDateTime = @orderDateTime And userID = @userID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@orderDateTime", dateTimePayment)
                MySqlCommand.Parameters.AddWithValue("@userID", userID)
                Dim orderID = MySqlCommand.ExecuteScalar()

                Dim bookIDBuyNow As String = qsBookID
                Dim quantityBuyNow As String = qsQuantity

                'If String.IsNullOrEmpty(bookIDBuyNow) Or String.IsNullOrEmpty(quantityBuyNow) Then
                'update book table

                Dim bookNum As Integer
                strSql = "Select COUNT(bookID) From Cart Where userID = @userID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@userID", userID)
                bookNum = Integer.Parse(MySqlCommand.ExecuteScalar())

                Dim bookIDArr(bookNum - 1) As Integer

                For i = 1 To bookNum
                    strSql = "Select bookID From ( " &
                        "Select bookID, ROW_NUMBER() OVER(ORDER BY bookID) AS ROW " &
                        "FROM Cart Where userID = @userID ) As TMP " &
                        "Where ROW = @row"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@userID", userID)
                    MySqlCommand.Parameters.AddWithValue("@row", i)
                    bookIDArr(i - 1) = MySqlCommand.ExecuteScalar()
                Next

                For i = 0 To bookNum - 1
                    Dim quantity As Integer
                    Dim cartNum As Integer
                    strSql = "Select bookQuantity From Book Where bookID = @bookID"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    quantity = Integer.Parse(MySqlCommand.ExecuteScalar())

                    strSql = "Select cartNum From Book Inner Join Cart On Cart.bookID = Book.bookID And Cart.bookID = @bookID"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    cartNum = Integer.Parse(MySqlCommand.ExecuteScalar())

                    quantity = quantity - cartNum

                    strSql = "Update Book Set bookQuantity = @bookQuantity Where bookID = @bookID"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    MySqlCommand.Parameters.AddWithValue("@bookQuantity", quantity)
                    MySqlCommand.ExecuteNonQuery()



                    Dim sales As Integer
                    strSql = "Select bookSales From Book Where bookID = @bookID"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    sales = Integer.Parse(MySqlCommand.ExecuteScalar())

                    sales = sales + cartNum

                    strSql = "Update Book Set bookSales = @bookSales Where bookID = @bookID"
                    MySqlCommand = New SqlCommand(strSql, db.conn)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    MySqlCommand.Parameters.AddWithValue("@bookSales", sales)
                    MySqlCommand.ExecuteNonQuery()

                    'update orderdetail table

                    strSql = "INSERT into OrderDetail(orderID, bookID, odQuantity) " +
                    "VALUES (@orderID, @bookID, @odQuantity)"

                    MySqlCommand = New SqlCommand(strSql, db.conn)

                    MySqlCommand.Parameters.AddWithValue("@orderID", orderID)
                    MySqlCommand.Parameters.AddWithValue("@bookID", bookIDArr(i))
                    MySqlCommand.Parameters.AddWithValue("@odQuantity", cartNum)

                    MySqlCommand.ExecuteNonQuery()



                Next

                'End If
                db.CloseConnection()
            End If

        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Public Function UpdateCart(userID As Integer) As Boolean
        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then

                strSql = "DELETE FROM Cart WHERE userID = @userID"

                MySqlCommand = New SqlCommand(strSql, db.conn)

                MySqlCommand.Parameters.AddWithValue("@userID", userID)

                MySqlCommand.ExecuteNonQuery()

                db.CloseConnection()
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Public Function UpdateBookBN(orderDetail As ModelLibrary.OrderDetail) As Boolean
        'userID As Integer, bookIDBuyNow As Integer, quantityBuyNow As Integer
        Dim userID As Integer = orderDetail.userID
        Dim bookIDBuyNow As Integer = orderDetail.bookID
        Dim quantityBuyNow As Integer = orderDetail.odQuantity
        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then

                'update book table
                Dim quantity As Integer
                strSql = "Select bookQuantity From Book Where bookID = @bookID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@bookID", bookIDBuyNow)
                quantity = Integer.Parse(MySqlCommand.ExecuteScalar())

                quantity = quantity - quantityBuyNow

                strSql = "Update Book Set bookQuantity = @bookQuantity Where bookID = @bookID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@bookID", bookIDBuyNow)
                MySqlCommand.Parameters.AddWithValue("@bookQuantity", quantity)
                MySqlCommand.ExecuteNonQuery()

                Dim sales As Integer
                strSql = "Select bookSales From Book Where bookID = @bookID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@bookID", bookIDBuyNow)
                sales = Integer.Parse(MySqlCommand.ExecuteScalar())

                sales = sales + quantityBuyNow

                strSql = "Update Book Set bookSales = @bookSales Where bookID = @bookID"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@bookID", bookIDBuyNow)
                MySqlCommand.Parameters.AddWithValue("@bookSales", sales)
                MySqlCommand.ExecuteNonQuery()

                db.CloseConnection()
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function


    Public Function InsertOrderDetail(orderDetail As ModelLibrary.OrderDetail) As Boolean
        'userID As Integer, dateTimePayment As DateTime, bookIDBuyNow As Integer,
        'quantityBuyNow As Integer
        Dim dateTimePayment As DateTime = orderDetail.orderDateTime
        Dim userID As Integer = orderDetail.userID
        Dim bookIDBuyNow As Integer = orderDetail.bookID
        Dim quantityBuyNow As Integer = orderDetail.odQuantity
        Try
            Dim MySqlCommand As New SqlCommand
            Dim strSql As String

            If db.OpenConnection() = True Then

                strSql = "select orderid from bookorder where orderdatetime = @orderdatetime and userid = @userid"
                MySqlCommand = New SqlCommand(strSql, db.conn)
                MySqlCommand.Parameters.AddWithValue("@orderdatetime", dateTimePayment)
                MySqlCommand.Parameters.AddWithValue("@userid", userID)
                Dim orderid = MySqlCommand.ExecuteScalar()

                strSql = "INSERT into OrderDetail(orderID, bookID, odQuantity) " +
                "VALUES (@orderID, @bookID, @odQuantity)"

                MySqlCommand = New SqlCommand(strSql, db.conn)

                MySqlCommand.Parameters.AddWithValue("@orderID", orderid)
                MySqlCommand.Parameters.AddWithValue("@bookID", bookIDBuyNow)
                MySqlCommand.Parameters.AddWithValue("@odQuantity", quantityBuyNow)

                MySqlCommand.ExecuteNonQuery()

                db.CloseConnection()
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Public Function SearchBookCartResult(userID As Integer) As List(Of ModelLibrary.BookCartCombined)
        Dim result As New List(Of ModelLibrary.BookCartCombined)

        If db.OpenConnection = True Then
            Dim strSqlCart = "Select * From Book Inner Join Cart " &
               "On Book.bookID = Cart.bookID " &
                "And Cart.userID = @userID " &
                "And cartNum > 0"
            Dim searchCmd As New SqlCommand
            searchCmd = New SqlCommand(strSqlCart, db.conn)
            searchCmd.Parameters.AddWithValue("@userID", userID)
            Dim da As New SqlDataAdapter(searchCmd)
            Dim ds As DataSet = New DataSet()
            ds.Clear()
            da.Fill(ds, "Cart")

            If ds.Tables("Cart").Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables("Cart").Rows
                    Dim item As New ModelLibrary.BookCartCombined()
                    item.bookID = Convert.ToInt32(row("bookID"))
                    item.bookTitle = row("bookTitle").ToString()
                    item.bookAuthor = row("bookAuthor").ToString()
                    item.bookImg = row("bookImg").ToString()
                    item.bookCategory = row("bookCategory").ToString()
                    item.bookPrice = Decimal.Parse(row("bookPrice").ToString())
                    item.bookPublishedDateTime = DateTime.Parse(row("bookPublishedDateTime").ToString)
                    item.bookSales = Convert.ToInt32(row("bookSales").ToString())
                    item.bookStatus = row("bookStatus").ToString()
                    item.bookQuantity = Convert.ToInt32(row("bookQuantity").ToString())
                    item.bookPublisher = row("bookPublisher").ToString()
                    item.bookWeight = Decimal.Parse(row("bookWeight").ToString())
                    item.bookDes = row("bookDes").ToString()
                    item.userID = Convert.ToInt32(row("userID").ToString())
                    item.cartNum = Convert.ToInt32(row("cartNum").ToString())
                    result.Add(item)
                Next
            End If


            db.CloseConnection()

        End If

        Return result
    End Function

    Public Function SearchBookBNResult(userID As Integer, bookID As String, quantity As Integer) As List(Of ModelLibrary.BookCartCombined)
        Dim result As New List(Of ModelLibrary.BookCartCombined)

        If db.OpenConnection = True Then
            Dim strSqlCart = "Select * From Book " &
               " WHERE " &
                "bookID = @bookID"
            Dim searchCmd As New SqlCommand
            searchCmd = New SqlCommand(strSqlCart, db.conn)
            'searchCmd.Parameters.AddWithValue("@userID", Session("UserID"))
            searchCmd.Parameters.AddWithValue("@bookID", bookID)
            Dim da As New SqlDataAdapter(searchCmd)
            Dim ds As New DataSet()
            'da.SelectCommand = searchCmd
            ds.Clear()
            da.Fill(ds, "Book")

            If ds.Tables("Book").Rows.Count > 0 Then
                Dim row As DataRow = ds.Tables("Book").Rows(0)

                Dim item As New ModelLibrary.BookCartCombined()
                item.bookID = Convert.ToInt32(row("bookID"))
                item.bookTitle = row("bookTitle").ToString()
                item.bookAuthor = row("bookAuthor").ToString()
                item.bookImg = row("bookImg").ToString()
                item.bookCategory = row("bookCategory").ToString()
                item.bookPrice = Decimal.Parse(row("bookPrice").ToString())
                result.Add(item)
            End If

            db.CloseConnection()
        End If
        Return result
    End Function

    Public Function SearchEmailPaymentPage(userID As Integer) As String
        Dim dbEmail As String
        If db.OpenConnection = True Then

            Dim strSqlUser = "Select userEmail From UserAcc Where userID = @userID"
            Dim searchCmd As New SqlCommand
            searchCmd = New SqlCommand(strSqlUser, db.conn)
            searchCmd.Parameters.AddWithValue("userID", userID)

            dbEmail = searchCmd.ExecuteScalar()

            db.CloseConnection()
        Else
            dbEmail = String.Empty
        End If
        Return dbEmail
    End Function

    Public Function SearchCartResult(userID As Integer) As List(Of ModelLibrary.BookCartCombined)
        Dim items As New List(Of ModelLibrary.BookCartCombined)()

        If db.OpenConnection = True Then
            Dim strSqlCart = "Select * From Book Inner Join Cart " &
               "On Book.bookID = Cart.bookID " &
                "And Cart.userID = @userID " &
                "And cartNum > 0"
            Dim searchCmd As New SqlCommand
            searchCmd = New SqlCommand(strSqlCart, db.conn)
            searchCmd.Parameters.AddWithValue("@userID", userID)
            Dim da As New SqlDataAdapter(searchCmd)
            Dim ds As DataSet = New DataSet()
            ds.Clear()
            da.Fill(ds, "Cart")

            If ds.Tables("Cart").Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables("Cart").Rows
                    Dim item As New ModelLibrary.BookCartCombined()
                    item.bookID = Convert.ToInt32(row("bookID"))
                    item.bookTitle = row("bookTitle").ToString()
                    item.bookAuthor = row("bookAuthor").ToString()
                    item.bookImg = row("bookImg").ToString()
                    item.bookCategory = row("bookCategory").ToString()
                    item.bookPrice = Decimal.Parse(row("bookPrice").ToString())
                    item.bookPublishedDateTime = DateTime.Parse(row("bookPublishedDateTime").ToString)
                    item.bookSales = Convert.ToInt32(row("bookSales").ToString())
                    item.bookStatus = row("bookStatus").ToString()
                    item.bookQuantity = Convert.ToInt32(row("bookQuantity").ToString())
                    item.bookPublisher = row("bookPublisher").ToString()
                    item.bookWeight = Decimal.Parse(row("bookWeight").ToString())
                    item.bookDes = row("bookDes").ToString()
                    item.userID = Convert.ToInt32(row("userID").ToString())
                    item.cartNum = Convert.ToInt32(row("cartNum").ToString())
                    items.Add(item)
                Next


            End If
            db.CloseConnection()

        End If
        Return items
    End Function

    Public Function DeleteCartItem(userID As Integer, bookID As Integer) As Boolean
        Dim strSql As String
        Try
            If db.OpenConnection = True Then
                Dim deleteCmd As New SqlCommand
                strSql = "DELETE FROM Cart WHERE bookID = @bookID AND userID = @userID"
                deleteCmd = New SqlCommand(strSql, db.conn)
                deleteCmd.Parameters.AddWithValue("@bookID", bookID)
                deleteCmd.Parameters.AddWithValue("@userID", userID)
                deleteCmd.ExecuteNonQuery()
                db.CloseConnection()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function UpdateCartNum(userID As Integer, quantity As Integer, bookID As Integer) As Boolean
        If db.OpenConnection() Then
            Dim strSqlUpdate = "UPDATE Cart SET cartNum = @cartNum WHERE bookID = @bookID and userID = @userID"
            Dim updateCmd As New SqlCommand(strSqlUpdate, db.conn)
            updateCmd.Parameters.AddWithValue("@cartNum", quantity)
            updateCmd.Parameters.AddWithValue("@bookID", bookID)
            updateCmd.Parameters.AddWithValue("@userID", userID)
            updateCmd.ExecuteNonQuery()
            db.CloseConnection()
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SearchBookQuantity(bookID As Integer) As Integer
        Dim bookQuantity As Integer
        Dim searchCmd As New SqlCommand
        Dim strSql As String
        If db.OpenConnection = True Then
            strSql = "Select bookQuantity from Book Where bookID = @bookID"
            searchCmd = New SqlCommand(strSql, db.conn)
            searchCmd.Parameters.AddWithValue("@bookID", bookID)

            bookQuantity = Integer.Parse(searchCmd.ExecuteScalar())

            db.CloseConnection()
            Return bookQuantity
        Else
            Return 0
        End If
    End Function

    Public Function SearchWishlistResult(userID As String) As List(Of ModelLibrary.BookWishlistCombined)
        Dim items As New List(Of ModelLibrary.BookWishlistCombined)()

        If db.OpenConnection = True Then
            Dim strSqlWishlist = "Select * From Book Inner Join Wishlist " &
               "On Book.bookID = Wishlist.bookID " &
                "And Wishlist.userID = @userID " &
                "And wishlistPreference = 1"
            Dim searchCmd As New SqlCommand
            searchCmd = New SqlCommand(strSqlWishlist, db.conn)
            searchCmd.Parameters.AddWithValue("@userID", userID)
            Dim da As New SqlDataAdapter(searchCmd)
            Dim ds As DataSet = New DataSet()
            ds.Clear()
            da.Fill(ds, "Wishlist")

            If ds.Tables("Wishlist").Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables("Wishlist").Rows
                    Dim item As New ModelLibrary.BookWishlistCombined()
                    item.bookID = Convert.ToInt32(row("bookID"))
                    item.bookTitle = row("bookTitle").ToString()
                    item.bookAuthor = row("bookAuthor").ToString()
                    item.bookImg = row("bookImg").ToString()
                    item.bookCategory = row("bookCategory").ToString()
                    item.bookPrice = Decimal.Parse(row("bookPrice").ToString())
                    item.bookPublishedDateTime = DateTime.Parse(row("bookPublishedDateTime").ToString)
                    item.bookSales = Convert.ToInt32(row("bookSales").ToString())
                    item.bookStatus = row("bookStatus").ToString()
                    item.bookQuantity = Convert.ToInt32(row("bookQuantity").ToString())
                    item.bookPublisher = row("bookPublisher").ToString()
                    item.bookWeight = Decimal.Parse(row("bookWeight").ToString())
                    item.bookDes = row("bookDes").ToString()
                    item.userID = Convert.ToInt32(row("userID").ToString())
                    item.wishlistPreference = Convert.ToInt32(row("wishlistPreference").ToString())
                    items.Add(item)
                Next


            End If
            db.CloseConnection()


        End If

        Return items
    End Function

    Public Function UpdateNullWishlistItem(userID As Integer, bookID As Integer) As Integer
        Dim value As Integer
        If db.OpenConnection = True Then
            Dim strUpdate = "Update Wishlist SET wishlistPreference = @wishlistPreference Where userID = @userID and bookID = @bookID"

            Dim updateCmd As New SqlCommand
            updateCmd = New SqlCommand(strUpdate, db.conn)
            updateCmd.Parameters.AddWithValue("@userID", userID)
            updateCmd.Parameters.AddWithValue("@bookID", bookID)
            updateCmd.Parameters.AddWithValue("@wishlistPreference", 0)


            Dim intInsertStatus As Integer = updateCmd.ExecuteNonQuery()

        End If

        db.CloseConnection()
        Return value
    End Function

End Class
