' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
