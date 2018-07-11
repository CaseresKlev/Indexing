Imports MySql.Data.MySqlClient

Public Class db_config
    Public con As New MySqlConnection("server= localhost; user id = 'root'; password= '';database=db_rrms;")
    'Public con As New MySqlConnection("server= 192.168.0.101; user id = 'klevie'; password= 'klevie';database=db_sboam;")

    Sub New()

    End Sub



    Public Function getCon()
        Return con
    End Function
End Class
