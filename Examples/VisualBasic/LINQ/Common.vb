
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace LINQ
    Class Common

        ''' <summary>
        ''' Return first manager from Managers which is an enumeration of instances of the Manager class. 
        ''' </summary>
        Public Shared Function GetManager() As Manager
            Dim managers As IEnumerator(Of Manager) = GetManagers().GetEnumerator()
            managers.MoveNext()
            Return managers.Current
        End Function

        ''' <summary>
        ''' Return an enumeration of instances of the Client class. 
        ''' </summary>
        Public Shared Function GetClients() As IEnumerable(Of Client)
            Dim returnCollection As New List(Of Client)()
            For Each manager As Manager In GetManagers()
                For Each contract As Contract In manager.Contracts
                    returnCollection.Add(contract.Client)
                Next
            Next
            Return returnCollection
        End Function

        ''' <summary>
        '''  Return an enumeration of instances of the Manager class.
        ''' </summary>
        Public Shared Function GetManagers() As IEnumerable(Of Manager)
            Dim returnCollection As New List(Of Manager)()
            Dim manager As New Manager() With { _
                .Name = "John Smith", _
                .Age = 36, _
                .Photo = Photo() _
            }
            manager.Contracts = New Contract() {New Contract() With { _
                .Client = New Client() With { _
                    .Name = "A Company" _
                }, _
                .Manager = manager, _
                .Price = 1200000, _
                .[Date] = New DateTime(2015, 1, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "B Ltd." _
                }, _
                .Manager = manager, _
                .Price = 750000, _
                .[Date] = New DateTime(2015, 4, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "C & D" _
                }, _
                .Manager = manager, _
                .Price = 350000, _
                .[Date] = New DateTime(2015, 7, 1) _
            }}

            returnCollection.Add(manager)


            manager = New Manager() With { _
                .Name = "Tony Anderson", _
                .Age = 37, _
                .Photo = Photo() _
            }
            manager.Contracts = New Contract() {New Contract() With { _
                .Client = New Client() With { _
                    .Name = "E Corp." _
                }, _
                .Manager = manager, _
                .Price = 650000, _
                .[Date] = New DateTime(2015, 2, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "F & Partners" _
                }, _
                .Manager = manager, _
                .Price = 550000, _
                .[Date] = New DateTime(2015, 8, 1) _
            }}

            returnCollection.Add(manager)

            manager = New Manager() With { _
                .Name = "July James", _
                .Age = 38, _
                .Photo = Photo() _
            }
            manager.Contracts = New Contract() {New Contract() With { _
                .Client = New Client() With { _
                    .Name = "G & Co." _
                }, _
                .Manager = manager, _
                .Price = 350000, _
                .[Date] = New DateTime(2015, 2, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "H Group" _
                }, _
                .Manager = manager, _
                .Price = 250000, _
                .[Date] = New DateTime(2015, 5, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "I & Sons" _
                }, _
                .Manager = manager, _
                .Price = 100000, _
                .[Date] = New DateTime(2015, 7, 1) _
            }, New Contract() With { _
                .Client = New Client() With { _
                    .Name = "J Ent." _
                }, _
                .Manager = manager, _
                .Price = 100000, _
                .[Date] = New DateTime(2015, 8, 1) _
            }}
            returnCollection.Add(manager)

            Return returnCollection
        End Function

        ''' <summary>
        ''' Return an array of photo bytes. 
        ''' </summary>
        Private Shared Function Photo() As Byte()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            ' Load the photo and read all bytes.
            Dim imgdata As Byte() = System.IO.File.ReadAllBytes(dataDir & Convert.ToString("photo.png"))
            Return imgdata
        End Function

        ''' <summary>
        '''  Return an enumeration of instances of the Contract class.
        ''' </summary>
        Public Shared Function GetContracts() As IEnumerable(Of Contract)
            Dim returnCollection As New List(Of Contract)()
            For Each manager As Manager In GetManagers()
                For Each contract As Contract In manager.Contracts
                    'yield Return contract
                    returnCollection.Add(contract)
                Next
            Next
            Return returnCollection
        End Function


    End Class
End Namespace

